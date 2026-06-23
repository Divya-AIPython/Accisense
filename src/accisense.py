import argparse
import random
import re
import sys
import threading
import time
import cv2
import requests
from ultralytics import YOLO
import easyocr

UK_PLATE_FULL = re.compile(r"^[A-Z]{2}\d{2}[A-Z]{3}$")

VEHICLES = {"car", "bus", "truck", "motorcycle"}
ALERT_URL = "http://127.0.0.1:8000/alert"
DEFAULT_ACK_URL = "http://127.0.0.1:8000/ack"
NOT_READABLE = "NOT READABLE"


class AccidentSharedState:
    """Thread-safe plates + box list for video overlay."""

    def __init__(self):
        self.lock = threading.Lock()
        self.v1 = NOT_READABLE
        self.v2 = NOT_READABLE
        self.plate_boxes = []

    def set_plates(self, v1, v2, boxes):
        with self.lock:
            self.v1, self.v2 = v1, v2
            self.plate_boxes = list(boxes)

    def snapshot(self):
        with self.lock:
            return self.v1, self.v2, list(self.plate_boxes)


def hospital_acknowledged(ack_url):
    """
    True when hospital dashboard has marked the case accepted.
    Expects HTTP 200 and JSON {"accepted": true} or {"hospital_accepted": true},
    or plain text containing 'accepted'.
    """
    try:
        r = requests.get(ack_url, timeout=(0.25, 1.0))
        if r.status_code != 200:
            return False
        try:
            j = r.json()
            if j.get("accepted") in (True, "true", 1, "1"):
                return True
            if j.get("hospital_accepted") in (True, "true", 1, "1"):
                return True
            if j.get("acknowledged") in (True, "true", 1, "1"):
                return True
        except ValueError:
            pass
        return "accepted" in r.text.lower()
    except Exception:
        return False


def send_alert_async(p1, p2, location, alert_url):
    """Never blocks the video loop; short connect timeout."""

    def _post():
        payload = {"vehicle_1": p1, "vehicle_2": p2, "location": location}
        try:
            requests.post(
                alert_url,
                json=payload,
                timeout=(0.35, 2.0),
            )
        except Exception:
            pass

    threading.Thread(target=_post, daemon=True).start()


def clean_text(s):
    s = (s or "").upper()
    s = re.sub(r"[^A-Z0-9]", "", s)
    return s


def _fix_digit_char(c):
    m = {"O": "0", "Q": "0", "D": "0", "Z": "7", "I": "1", "L": "1", "S": "5", "B": "8", "G": "6"}
    if c.isdigit():
        return c
    return m.get(c, c)


def normalize_plate_candidate(s):
    s = clean_text(s)
    if not s:
        return None
    if len(s) >= 7:
        candidates = [s[-7:], s[:7]]
    else:
        candidates = [s]
    best = None
    for c in candidates:
        if len(c) != 7:
            continue
        area, dig_raw, tail_raw = c[:2], c[2:4], c[4:]
        dig = "".join(_fix_digit_char(x) for x in dig_raw)
        tail = re.sub(r"[^A-Z]", "", tail_raw)
        if len(tail) != 3 or not area.isalpha() or not dig.isdigit() or len(dig) != 2:
            continue
        cand = area + dig + tail
        if UK_PLATE_FULL.match(cand):
            best = cand
            break
    return best


def box_iou(a, b):
    ax1, ay1, ax2, ay2 = a
    bx1, by1, bx2, by2 = b
    x1, y1 = max(ax1, bx1), max(ay1, by1)
    x2, y2 = min(ax2, bx2), min(ay2, by2)
    iw, ih = max(0, x2 - x1), max(0, y2 - y1)
    inter = iw * ih
    if inter == 0:
        return 0.0
    area_a = (ax2 - ax1) * (ay2 - ay1)
    area_b = (bx2 - bx1) * (by2 - by1)
    return inter / float(area_a + area_b - inter)


def box_center(b):
    x1, y1, x2, y2 = b
    return ((x1 + x2) * 0.5, (y1 + y2) * 0.5)


def plate_assign_vehicle(plate_box, vehicles_sorted):
    if not vehicles_sorted:
        return None
    cx, cy = box_center(plate_box)
    for i, v in enumerate(vehicles_sorted):
        x1, y1, x2, y2 = v
        if x1 <= cx <= x2 and y1 <= cy <= y2:
            return min(i, 1)
    best_i, best_d = 0, 1e18
    for i, v in enumerate(vehicles_sorted[:2]):
        vx = (v[0] + v[2]) * 0.5
        d = abs(cx - vx)
        if d < best_d:
            best_d, best_i = d, i
    return best_i


def preprocess_for_ocr(crop_bgr, fast=True):
    """fast=True: fewer variants so plates update faster."""
    gray = cv2.cvtColor(crop_bgr, cv2.COLOR_BGR2GRAY)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    if fast:
        return [gray, clahe.apply(gray)]
    out = [gray, clahe.apply(gray)]
    blur = cv2.GaussianBlur(gray, (3, 3), 0)
    out.append(
        cv2.adaptiveThreshold(
            blur, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 2
        )
    )
    return out


def ocr_plate(crop, reader, min_conf=0.45, fast=True):
    """
    Only accept strict UK current-style plates (2 letters + 2 digits + 3 letters).
    """
    best_plate, best_cf = None, 0.0
    scales = (2.0,) if fast else (1.6, 2.2)
    for img in preprocess_for_ocr(crop, fast=fast):
        for scale in scales:
            h, w = img.shape[:2]
            resized = cv2.resize(
                img, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_CUBIC
            )
            det = reader.readtext(
                resized,
                detail=1,
                allowlist="ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789",
                paragraph=False,
            )
            det_sorted = sorted(det, key=lambda it: -float(it[2]) if len(it) > 2 else 0.0)
            for item in det_sorted:
                if len(item) < 3:
                    continue
                _, text, conf = item[0], item[1], float(item[2])
                if conf < min_conf:
                    continue
                candidate = normalize_plate_candidate(text)
                if not candidate or not UK_PLATE_FULL.match(candidate):
                    continue
                if (
                    best_plate is None
                    or conf > best_cf
                    or (conf == best_cf and candidate < best_plate)
                ):
                    best_cf, best_plate = conf, candidate
    return best_plate, best_cf


def majority(lst):
    if not lst:
        return NOT_READABLE
    return max(set(lst), key=lst.count)


def seed_all(seed=42):
    """Reduce run-to-run drift in YOLO / EasyOCR."""
    random.seed(seed)
    try:
        import numpy as np
        np.random.seed(seed)
    except ImportError:
        pass
    try:
        import torch
        torch.manual_seed(seed)
        if torch.cuda.is_available():
            torch.cuda.manual_seed_all(seed)
            torch.backends.cudnn.deterministic = True
            torch.backends.cudnn.benchmark = False
    except ImportError:
        pass


def plate_pipeline_worker(
    frame_bgr,
    vehicles_sorted,
    plate_model,
    reader,
    ocr_conf,
    shared,
    location,
    alert_url,
    plates_ready_event,
    send_alert,
):
    """Runs YOLO plate + EasyOCR off the main thread so video keeps updating."""
    votes1, votes2 = [], []
    plate_boxes_xyxy = []
    last_published = (NOT_READABLE, NOT_READABLE)

    def publish_if_changed(b1, b2):
        nonlocal last_published
        pair = (b1, b2)
        if pair == last_published:
            return
        last_published = pair
        shared.set_plates(b1, b2, list(plate_boxes_xyxy))
        plates_ready_event.set()

    pres = plate_model(frame_bgr, imgsz=320, conf=0.42, verbose=False)[0]
    ordered = []
    if pres.boxes is not None and len(pres.boxes) > 0:
        xy_t = pres.boxes.xyxy.cpu().numpy().tolist()
        cf_t = pres.boxes.conf.cpu().numpy().tolist()
        ordered = sorted(zip(cf_t, xy_t), key=lambda p: -p[0])

    for _, pb in ordered:
        x1, y1, x2, y2 = map(float, pb)
        x1, y1 = max(0, int(x1 - 4)), max(0, int(y1 - 4))
        x2, y2 = min(frame_bgr.shape[1], int(x2 + 4)), min(frame_bgr.shape[0], int(y2 + 4))
        plate_boxes_xyxy.append((x1, y1, x2, y2))
        crop = frame_bgr[y1:y2, x1:x2]
        if crop.size == 0 or len(vehicles_sorted) < 1:
            continue
        plate_txt, _ = ocr_plate(crop, reader, min_conf=ocr_conf, fast=True)
        if not plate_txt:
            continue
        slot = plate_assign_vehicle((x1, y1, x2, y2), vehicles_sorted)
        if slot == 0:
            votes1.append(plate_txt)
        elif slot == 1:
            votes2.append(plate_txt)
        else:
            if len(votes1) <= len(votes2):
                votes1.append(plate_txt)
            else:
                votes2.append(plate_txt)
        best1 = majority(votes1) if votes1 else NOT_READABLE
        best2 = majority(votes2) if votes2 else NOT_READABLE
        publish_if_changed(best1, best2)

    best1 = majority(votes1) if votes1 else NOT_READABLE
    best2 = majority(votes2) if votes2 else NOT_READABLE
    publish_if_changed(best1, best2)

    print(f"Vehicle 1 plate: {best1}")
    print(f"Vehicle 2 plate: {best2}")
    print("===================================\n")
    sys.stdout.flush()

    if send_alert:
        send_alert_async(best1, best2, location, alert_url)


def draw_boxes(frame, vehicles, accident_detected):
    color = (0, 0, 255) if accident_detected else (0, 255, 0)
    thickness = 3
    for box in vehicles:
        x1, y1, x2, y2 = map(int, box)
        cv2.rectangle(frame, (x1, y1), (x2, y2), color, thickness)


def draw_plate_boxes(frame, plate_xyxy_list):
    for (x1, y1, x2, y2) in plate_xyxy_list:
        cv2.rectangle(
            frame,
            (int(x1), int(y1)),
            (int(x2), int(y2)),
            (0, 0, 255),
            2,
        )


def draw_accident_banner(frame):
    h, w = frame.shape[:2]
    overlay = frame.copy()
    cv2.rectangle(overlay, (0, 0), (w, 80), (0, 0, 200), -1)
    cv2.addWeighted(overlay, 0.6, frame, 0.4, 0, frame)
    text = "ACCIDENT DETECTED!"
    font = cv2.FONT_HERSHEY_DUPLEX
    scale, thick = 1.6, 3
    tw, _ = cv2.getTextSize(text, font, scale, thick)[0]
    cx = (w - tw) // 2
    cv2.putText(frame, text, (cx + 2, 56), font, scale, (0, 0, 0), thick + 2)
    cv2.putText(frame, text, (cx, 54), font, scale, (255, 255, 255), thick)


def main():
    parser = argparse.ArgumentParser(description="AcciSense – accident detector")
    parser.add_argument("--source", required=True, help="Video path or camera index")
    parser.add_argument("--accident_model", required=True, help="YOLO accident/vehicle model")
    parser.add_argument("--plate_model", required=True, help="YOLO plate model")
    parser.add_argument("--location", default="Junction Camera 4")
    parser.add_argument("--alert_url", default=ALERT_URL, help="POST JSON alert endpoint")
    parser.add_argument(
        "--ack_url",
        default=DEFAULT_ACK_URL,
        help="GET this until JSON accepted=true (hospital acknowledged)",
    )
    parser.add_argument("--iou_thresh", type=float, default=0.08)
    parser.add_argument(
        "--confirm_frames",
        type=int,
        default=1,
        help="Consecutive frames with vehicle overlap to confirm accident (1 = fastest)",
    )
    parser.add_argument("--ocr_conf", type=float, default=0.40, help="Min EasyOCR confidence")
    parser.add_argument(
        "--seed",
        type=int,
        default=42,
        help="RNG seed for more stable YOLO/EasyOCR across repeated runs",
    )
    args = parser.parse_args()

    seed_all(args.seed)

    acc_model = YOLO(args.accident_model)
    plate_model = YOLO(args.plate_model)
    reader = easyocr.Reader(["en"], gpu=False)

    source = int(args.source) if args.source.isdigit() else args.source
    cap = cv2.VideoCapture(source)
    if not cap.isOpened():
        print("ERROR: Cannot open video source")
        sys.exit(1)
    cap.set(cv2.CAP_PROP_POS_FRAMES, 0)

    print("\n--- AcciSense STARTED ---\n")

    accident_detected = False
    touch_counter = 0

    shared = AccidentSharedState()
    plates_ready = threading.Event()

    # ── Rolling best-frame buffer ─────────────────────────────────────────────
    # Keep track of the sharpest pre-collision frame so the plate is read
    # before it becomes occluded or motion-blurred during overlap.
    best_frame_snap = None
    best_frame_veh  = None
    best_frame_lap  = -1.0
    early_plate_started = False
    # ─────────────────────────────────────────────────────────────────────────

    while True:
        ret, frame = cap.read()
        if not ret:
            break

        results = acc_model(frame, imgsz=480, conf=0.4, verbose=False)[0]
        vehicles = []

        if results.boxes is not None:
            for box, cls, conf in zip(
                results.boxes.xyxy,
                results.boxes.cls,
                results.boxes.conf,
            ):
                name = results.names[int(cls)].lower()
                if name in VEHICLES and float(conf) > 0.4:
                    vehicles.append(box.cpu().numpy())

        vehicles.sort(key=lambda b: float(b[0]))

        # ── Update rolling best-frame buffer whenever >=2 vehicles visible ──
        if not accident_detected and len(vehicles) >= 2:
            gray_full = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            frame_lap  = cv2.Laplacian(gray_full, cv2.CV_64F).var()
            if frame_lap > best_frame_lap:
                best_frame_lap  = frame_lap
                best_frame_snap = frame.copy()
                best_frame_veh  = [v.copy() for v in vehicles]
        # ────────────────────────────────────────────────────────────────────

        if not accident_detected and len(vehicles) >= 2:
            best_iou = max(
                box_iou(vehicles[i], vehicles[j])
                for i in range(len(vehicles))
                for j in range(i + 1, len(vehicles))
            )

            if best_iou > args.iou_thresh:
                touch_counter += 1

                # ── Speculative OCR on first overlap using best buffered frame ──
                if touch_counter == 1 and not early_plate_started:
                    early_plate_started = True
                    ocr_snap = best_frame_snap if best_frame_snap is not None else frame.copy()
                    ocr_veh  = best_frame_veh  if best_frame_veh  is not None else [v.copy() for v in vehicles]
                    plates_ready.clear()
                # ──────────────────────────────────────────────────────────────

            else:
                touch_counter = 0
                early_plate_started = False

            if touch_counter >= args.confirm_frames:
                accident_detected = True
                print("===================================")
                print("ACCIDENT DETECTED")
                print("===================================")
                sys.stdout.flush()

                # Fallback: if speculative thread never launched, start now
                if not early_plate_started:
                    ocr_snap = best_frame_snap if best_frame_snap is not None else frame.copy()
                    ocr_veh  = best_frame_veh  if best_frame_veh  is not None else [v.copy() for v in vehicles]
                    plates_ready.clear()

                threading.Thread(
                    target=plate_pipeline_worker,
                    args=(
                        ocr_snap,
                        ocr_veh,
                        plate_model,
                        reader,
                        args.ocr_conf,
                        shared,
                        args.location,
                        args.alert_url,
                        plates_ready,
                        True,
                    ),
                    daemon=True,
                ).start()

        _, _, plate_boxes_draw = shared.snapshot()

        if accident_detected:
            draw_boxes(frame, vehicles, True)
            draw_accident_banner(frame)
            if plate_boxes_draw:
                draw_plate_boxes(frame, plate_boxes_draw)
        else:
            draw_boxes(frame, vehicles, False)

        cv2.imshow("AcciSense", frame)
        if cv2.waitKey(1) & 0xFF == 27:
            break

    cap.release()
    cv2.destroyAllWindows()
    print("\n--- AcciSense FINISHED ---\n")


if __name__ == "__main__":
    main()