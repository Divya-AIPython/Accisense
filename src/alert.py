from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, RedirectResponse, JSONResponse
from pydantic import BaseModel
from typing import Optional
import random
import threading
import time
import json
from datetime import datetime
import os

app = FastAPI(title="AcciSense Emergency Dashboard")

# ─────────────────────────────────────────────
# GLOBAL VOICE CONTROL
# ─────────────────────────────────────────────
voice_thread = None
stop_voice_flag = False

def voice_loop():
    global stop_voice_flag
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        speaker = win32com.client.Dispatch("SAPI.SpVoice")
        while not stop_voice_flag:
            speaker.Speak("Emergency! Accident Detected! Emergency!")
            time.sleep(2)
        pythoncom.CoUninitialize()
    except Exception as e:
        while not stop_voice_flag:
            print("🔊 VOICE ALERT: Emergency! Accident Detected!")
            time.sleep(2)

def start_voice_alert():
    global voice_thread, stop_voice_flag
    stop_voice_flag = False
    voice_thread = threading.Thread(target=voice_loop, daemon=True)
    voice_thread.start()

def stop_voice_alert():
    global stop_voice_flag
    stop_voice_flag = True

# ─────────────────────────────────────────────
# DATA
# ─────────────────────────────────────────────
POLICE_STATIONS = [
    "Central Police Station",
    "North Traffic Police HQ",
    "South City Police Station",
    "East Zone Police Station",
    "West District Control Room",
    "Highway Patrol Unit Alpha",
    "Metro Traffic Command",
    "Airport Zone Police Station",
    "Industrial Area Police Post",
    "City Control Room Central"
]

HOSPITALS = [
    {"name": "Apollo Emergency Center", "distance": "1.2 km"},
    {"name": "City Care Hospital", "distance": "2.4 km"},
    {"name": "Metro Trauma Hospital", "distance": "3.1 km"},
    {"name": "Green Cross Medical", "distance": "4.8 km"},
    {"name": "National Emergency Hospital", "distance": "6.2 km"},
]

VEHICLE_TYPES = {
    "TN01": {"type": "Sedan", "icon": "🚗"},
    "TN02": {"type": "SUV", "icon": "🚙"},
    "TN03": {"type": "Truck", "icon": "🚛"},
    "MH":   {"type": "Motorcycle", "icon": "🏍️"},
    "KA":   {"type": "Auto-Rickshaw", "icon": "🛺"},
    "AP":   {"type": "Bus", "icon": "🚌"},
    "TS":   {"type": "Van", "icon": "🚐"},
    "DL":   {"type": "Hatchback", "icon": "🚘"},
    "GJ":   {"type": "Pickup Truck", "icon": "🛻"},
    "RJ":   {"type": "Minibus", "icon": "🚎"},
    "HR":   {"type": "Bicycle", "icon": "🚲"},
    "UP":   {"type": "Ambulance", "icon": "🚑"},
    "WB":   {"type": "Taxi", "icon": "🚕"},
    "MP":   {"type": "Police Vehicle", "icon": "🚓"},
    "PB":   {"type": "Fire Truck", "icon": "🚒"},
}

BED_STATUS = [
    {"ward": "Emergency ICU", "total": 20, "occupied": 14, "available": 6, "status": "warning"},
    {"ward": "General Ward A", "total": 50, "occupied": 32, "available": 18, "status": "good"},
    {"ward": "General Ward B", "total": 50, "occupied": 45, "available": 5, "status": "critical"},
    {"ward": "Trauma Unit", "total": 15, "occupied": 8, "available": 7, "status": "good"},
    {"ward": "Pediatric Ward", "total": 25, "occupied": 10, "available": 15, "status": "good"},
    {"ward": "Orthopedic Ward", "total": 30, "occupied": 22, "available": 8, "status": "warning"},
    {"ward": "Cardiac ICU", "total": 12, "occupied": 11, "available": 1, "status": "critical"},
    {"ward": "Neurology Ward", "total": 20, "occupied": 13, "available": 7, "status": "good"},
]

FLEET_STATUS = [
    {"id": "AMB-001", "type": "Advanced Life Support", "status": "available", "driver": "Rajan Kumar", "location": "Base Station", "last_trip": "10:15 AM"},
    {"id": "AMB-002", "type": "Basic Life Support", "status": "dispatched", "driver": "Suresh Patel", "location": "En Route - NH44", "last_trip": "11:30 AM"},
    {"id": "AMB-003", "type": "Patient Transport", "status": "available", "driver": "Anil Sharma", "location": "Base Station", "last_trip": "09:45 AM"},
    {"id": "AMB-004", "type": "Advanced Life Support", "status": "maintenance", "driver": "N/A", "location": "Service Bay", "last_trip": "08:00 AM"},
    {"id": "AMB-005", "type": "Neonatal Transport", "status": "available", "driver": "Priya Nair", "location": "Base Station", "last_trip": "07:30 AM"},
    {"id": "AMB-006", "type": "Basic Life Support", "status": "dispatched", "driver": "Vikram Singh", "location": "City Hospital", "last_trip": "11:45 AM"},
    {"id": "POL-001", "type": "Police Patrol", "status": "available", "driver": "Officer Mehta", "location": "Zone 3", "last_trip": "10:00 AM"},
    {"id": "POL-002", "type": "Traffic Control", "status": "dispatched", "driver": "Officer Das", "location": "Accident Site", "last_trip": "11:30 AM"},
]

# ─────────────────────────────────────────────
# STATE
# ─────────────────────────────────────────────
state = {
    "current_case": None,
    "hospital_index": 0,
    "selected_police": None,
    "case_status": "idle",
    "alert_time": None,
    "accept_time": None,
    "case_history": [],
    "hospital_stats": {h["name"]: {"accepted": 0, "declined": 0, "times": []} for h in HOSPITALS},
    "police_stats": {p: {"cases": 0} for p in POLICE_STATIONS},
    "total_cases": 0,
    "terminal_log": [],
}

def log(msg):
    timestamp = datetime.now().strftime("%H:%M:%S")
    entry = f"[{timestamp}] {msg}"
    state["terminal_log"].append(entry)
    print(entry)
    if len(state["terminal_log"]) > 100:
        state["terminal_log"] = state["terminal_log"][-100:]

# ─────────────────────────────────────────────
# MODELS
# ─────────────────────────────────────────────
class AccidentAlert(BaseModel):
    vehicle_1: str
    vehicle_2: str
    latitude: Optional[float] = None
    longitude: Optional[float] = None

# ─────────────────────────────────────────────
# HELPER
# ─────────────────────────────────────────────
def get_vehicle_info(plate: str):
    for prefix, info in VEHICLE_TYPES.items():
        if plate.upper().startswith(prefix):
            return info
    return {"type": "Unknown Vehicle", "icon": "🚗"}

def current_hospital():
    idx = state["hospital_index"]
    if idx >= len(HOSPITALS):
        return None
    return HOSPITALS[idx]

# ─────────────────────────────────────────────
# ALERT ENDPOINT
# ─────────────────────────────────────────────
@app.post("/alert")
def receive_alert(alert: AccidentAlert):
    state["current_case"] = alert.dict()
    state["hospital_index"] = 0
    state["selected_police"] = random.choice(POLICE_STATIONS)
    state["case_status"] = "pending"
    state["alert_time"] = datetime.now().isoformat()
    state["accept_time"] = None
    state["total_cases"] += 1

    police = state["selected_police"]
    hospital = HOSPITALS[0]["name"]
    state["police_stats"][police]["cases"] += 1

    v1info = get_vehicle_info(alert.vehicle_1)
    v2info = get_vehicle_info(alert.vehicle_2)

    log(f"🚨 ACCIDENT DETECTED — Vehicle 1: {alert.vehicle_1} ({v1info['type']}) | Vehicle 2: {alert.vehicle_2} ({v2info['type']})")
    log(f"📍 GPS: {alert.latitude}, {alert.longitude}")
    log(f"🚓 Alert sent to Police Station: {police}")
    log(f"🏥 Contacting Hospital: {hospital}")

    start_voice_alert()
    return {"message": "Alert received", "redirect": "/police"}

@app.post("/accept")
def accept_case():
    state["case_status"] = "ambulance_dispatched"
    state["accept_time"] = datetime.now().isoformat()
    stop_voice_alert()

    hosp = current_hospital()
    if hosp:
        hname = hosp["name"]
        alert_time = datetime.fromisoformat(state["alert_time"])
        accept_time = datetime.fromisoformat(state["accept_time"])
        response_secs = int((accept_time - alert_time).total_seconds())
        state["hospital_stats"][hname]["accepted"] += 1
        state["hospital_stats"][hname]["times"].append(response_secs)

        log(f"✅ {hname} ACCEPTED the case")
        log(f"🚑 AMBULANCE DISPATCHED from {hname}")
        log(f"⏱️ Response time: {response_secs} seconds")

        c = state["current_case"]
        v1info = get_vehicle_info(c["vehicle_1"])
        v2info = get_vehicle_info(c["vehicle_2"])

        case = {
            "date": datetime.now().strftime("%Y-%m-%d"),
            "time": datetime.now().strftime("%H:%M:%S"),
            "v1": c["vehicle_1"],
            "v2": c["vehicle_2"],
            "v1type": v1info["type"],
            "v2type": v2info["type"],
            "v1icon": v1info["icon"],
            "v2icon": v2info["icon"],
            "hospital": hname,
            "police": state["selected_police"],
            "status": "Dispatched",
            "response_time": response_secs
        }
        state["case_history"].append(case)

    return RedirectResponse(url="/hospital", status_code=303)

@app.post("/decline")
def decline_case():
    hosp = current_hospital()
    if hosp:
        hname = hosp["name"]
        state["hospital_stats"][hname]["declined"] += 1
        log(f"❌ {hname} DECLINED — Moving to next hospital...")

    state["hospital_index"] += 1
    next_hosp = current_hospital()

    if next_hosp is None:
        state["case_status"] = "no_hospital"
        stop_voice_alert()
        log("⛔ No hospitals available — All declined")
    else:
        log(f"🏥 Now contacting: {next_hosp['name']}")

    return RedirectResponse(url="/hospital", status_code=303)

# ─────────────────────────────────────────────
# API ENDPOINTS
# ─────────────────────────────────────────────
@app.get("/api/state")
def get_state():
    return JSONResponse(state)

@app.get("/api/logs")
def get_logs():
    return JSONResponse({"logs": state["terminal_log"]})

@app.get("/api/hospital-stats")
def hospital_stats():
    stats = []
    for h in HOSPITALS:
        name = h["name"]
        s = state["hospital_stats"][name]
        total = s["accepted"] + s["declined"]
        avg_time = int(sum(s["times"]) / len(s["times"])) if s["times"] else 0
        decline_pct = round((s["declined"] / total * 100), 1) if total > 0 else 0
        stats.append({
            "name": name,
            "accepted": s["accepted"],
            "declined": s["declined"],
            "total": total,
            "avg_time": avg_time,
            "decline_pct": decline_pct,
        })
    return JSONResponse(stats)

@app.get("/api/bed-status")
def get_bed_status():
    return JSONResponse(BED_STATUS)

@app.get("/api/fleet-status")
def get_fleet_status():
    return JSONResponse(FLEET_STATUS)

# ─────────────────────────────────────────────
# LOGIN PAGE
# ─────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
def login_page():
    return HTMLResponse(LOGIN_HTML)

@app.post("/login")
async def login(request: Request):
    form = await request.form()
    role = form.get("role")
    if role == "police":
        return RedirectResponse(url="/police", status_code=303)
    elif role == "hospital":
        return RedirectResponse(url="/hospital", status_code=303)
    return RedirectResponse(url="/", status_code=303)

@app.get("/police", response_class=HTMLResponse)
def police_dashboard():
    return HTMLResponse(POLICE_HTML)

@app.get("/hospital", response_class=HTMLResponse)
def hospital_dashboard():
    return HTMLResponse(HOSPITAL_HTML)

# ─────────────────────────────────────────────
# HTML TEMPLATES
# ─────────────────────────────────────────────

LOGIN_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AcciSense — Emergency Response System</title>
<link href="https://fonts.googleapis.com/css2?family=Sora:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<style>
*{margin:0;padding:0;box-sizing:border-box;}
:root{
  --bg:#080c14;
  --panel:#0d1421;
  --panel2:#111827;
  --border:rgba(255,255,255,0.07);
  --border-hover:rgba(99,179,237,0.35);
  --red:#f43f5e;
  --red-glow:rgba(244,63,94,0.18);
  --blue:#3b82f6;
  --blue-glow:rgba(59,130,246,0.18);
  --teal:#14b8a6;
  --text:#f1f5f9;
  --text2:#94a3b8;
  --text3:#475569;
  --surface:rgba(255,255,255,0.03);
}
html,body{height:100%;overflow:hidden;}
body{
  background:var(--bg);
  font-family:'Sora',sans-serif;
  color:var(--text);
  display:flex;
  align-items:center;
  justify-content:center;
  position:relative;
}

/* BACKGROUND */
.bg-canvas{position:fixed;inset:0;z-index:0;overflow:hidden;}
.bg-grid{
  position:absolute;inset:0;
  background-image:
    linear-gradient(rgba(59,130,246,0.04) 1px, transparent 1px),
    linear-gradient(90deg, rgba(59,130,246,0.04) 1px, transparent 1px);
  background-size:48px 48px;
}
.bg-radial{
  position:absolute;inset:0;
  background:
    radial-gradient(ellipse 60% 50% at 15% 50%, rgba(244,63,94,0.07) 0%, transparent 70%),
    radial-gradient(ellipse 50% 60% at 85% 50%, rgba(59,130,246,0.07) 0%, transparent 70%),
    radial-gradient(ellipse 40% 40% at 50% 10%, rgba(20,184,166,0.05) 0%, transparent 60%);
}
.bg-scan{
  position:absolute;top:0;left:0;right:0;
  height:2px;
  background:linear-gradient(90deg, transparent, rgba(59,130,246,0.6), rgba(20,184,166,0.4), transparent);
  animation:scan 6s ease-in-out infinite;
  opacity:0.6;
}
@keyframes scan{0%{top:-2px;opacity:0;}10%{opacity:0.6;}90%{opacity:0.6;}100%{top:100vh;opacity:0;}}

/* SIDE PANELS */
.side-info{position:fixed;top:0;bottom:0;width:220px;display:flex;flex-direction:column;justify-content:center;gap:12px;padding:40px 24px;z-index:1;}
.side-info.left{left:0;}
.side-info.right{right:0;}
.info-chip{
  background:rgba(255,255,255,0.02);
  border:1px solid var(--border);
  border-radius:10px;
  padding:12px 14px;
  animation:fadeUp 0.6s ease both;
}
.info-chip:nth-child(1){animation-delay:0.1s;}
.info-chip:nth-child(2){animation-delay:0.2s;}
.info-chip:nth-child(3){animation-delay:0.3s;}
.info-chip:nth-child(4){animation-delay:0.4s;}
.info-chip:nth-child(5){animation-delay:0.5s;}
.chip-label{font-size:9px;font-weight:600;letter-spacing:1.5px;color:var(--text3);text-transform:uppercase;margin-bottom:5px;}
.chip-value{font-family:'JetBrains Mono',monospace;font-size:12px;font-weight:500;color:var(--text2);}
.chip-value.green{color:#4ade80;}
.chip-value.red{color:#f87171;}
.chip-value.amber{color:#fbbf24;}
.chip-dot{display:inline-block;width:6px;height:6px;border-radius:50%;margin-right:5px;vertical-align:middle;}
.chip-dot.green{background:#4ade80;box-shadow:0 0 6px #4ade80;}
.chip-dot.red{background:#f87171;animation:blink 1s infinite;}
.chip-dot.amber{background:#fbbf24;}
@keyframes blink{0%,100%{opacity:1;}50%{opacity:0.3;}}
@keyframes fadeUp{from{opacity:0;transform:translateY(12px);}to{opacity:1;transform:translateY(0);}}

/* MAIN CARD */
.wrap{position:relative;z-index:2;width:100%;max-width:420px;padding:20px;animation:fadeUp 0.5s ease both;}

/* LOGO */
.logo-area{text-align:center;margin-bottom:28px;}
.logo-mark{
  display:inline-flex;align-items:center;justify-content:center;
  width:64px;height:64px;
  background:linear-gradient(135deg,#f43f5e 0%,#dc2626 40%,#b91c1c 100%);
  border-radius:18px;
  box-shadow:0 0 0 1px rgba(244,63,94,0.3), 0 20px 60px rgba(244,63,94,0.25), 0 0 80px rgba(244,63,94,0.08);
  font-size:28px;
  margin:0 auto 16px;
  position:relative;
  animation:logoIn 0.7s cubic-bezier(0.34,1.56,0.64,1) both;
}
.logo-mark::after{
  content:'';position:absolute;inset:-1px;border-radius:19px;
  background:linear-gradient(135deg,rgba(255,255,255,0.15),transparent 60%);
  pointer-events:none;
}
@keyframes logoIn{from{opacity:0;transform:scale(0.6) translateY(10px);}to{opacity:1;transform:scale(1) translateY(0);}}
.logo-wordmark{font-family:'Sora',sans-serif;font-size:24px;font-weight:800;letter-spacing:-0.5px;color:var(--text);}
.logo-wordmark span{color:var(--red);}
.logo-tagline{font-size:11px;font-weight:400;color:var(--text3);letter-spacing:1px;margin-top:5px;text-transform:uppercase;}

/* CARD */
.card{
  background:var(--panel);
  border:1px solid var(--border);
  border-radius:20px;
  padding:28px;
  box-shadow:0 32px 64px rgba(0,0,0,0.5), inset 0 1px 0 rgba(255,255,255,0.04);
  position:relative;overflow:hidden;
}
.card::before{
  content:'';position:absolute;top:0;left:0;right:0;height:1px;
  background:linear-gradient(90deg,transparent,rgba(244,63,94,0.4),rgba(59,130,246,0.4),transparent);
}

/* ROLE SELECT */
.role-label{font-size:10px;font-weight:600;letter-spacing:2px;color:var(--text3);text-transform:uppercase;text-align:center;margin-bottom:14px;}
.role-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:24px;}
.role-btn{
  background:var(--surface);
  border:1px solid var(--border);
  border-radius:14px;
  padding:16px 10px;
  cursor:pointer;
  text-align:center;
  transition:all 0.2s cubic-bezier(0.4,0,0.2,1);
  position:relative;overflow:hidden;
  color:var(--text2);
  user-select:none;
}
.role-btn::before{content:'';position:absolute;inset:0;opacity:0;transition:opacity 0.2s;}
.role-btn:hover{border-color:rgba(255,255,255,0.12);color:var(--text);}
.role-btn.selected-police{
  border-color:rgba(244,63,94,0.5);
  background:rgba(244,63,94,0.06);
  color:var(--text);
  box-shadow:0 0 0 1px rgba(244,63,94,0.15), 0 8px 24px rgba(244,63,94,0.1);
}
.role-btn.selected-hospital{
  border-color:rgba(59,130,246,0.5);
  background:rgba(59,130,246,0.06);
  color:var(--text);
  box-shadow:0 0 0 1px rgba(59,130,246,0.15), 0 8px 24px rgba(59,130,246,0.1);
}
.role-icon{font-size:26px;margin-bottom:8px;display:block;}
.role-name{font-size:13px;font-weight:700;letter-spacing:0.3px;}
.role-desc{font-size:10px;color:var(--text3);margin-top:3px;letter-spacing:0.3px;}
.selected-police .role-desc{color:rgba(244,63,94,0.7);}
.selected-hospital .role-desc{color:rgba(59,130,246,0.7);}

/* SELECTED INDICATOR */
.sel-check{
  position:absolute;top:8px;right:8px;
  width:16px;height:16px;border-radius:50%;
  display:flex;align-items:center;justify-content:center;
  font-size:8px;font-weight:700;
  opacity:0;transition:0.2s;
}
.selected-police .sel-check{background:var(--red);color:white;opacity:1;}
.selected-hospital .sel-check{background:var(--blue);color:white;opacity:1;}

/* DIVIDER */
.divider{height:1px;background:var(--border);margin-bottom:22px;}

/* FIELDS */
.field{margin-bottom:14px;}
.field label{
  display:flex;align-items:center;gap:6px;
  font-size:10px;font-weight:600;letter-spacing:1.5px;
  color:var(--text3);text-transform:uppercase;margin-bottom:8px;
}
.field-ic{font-size:11px;opacity:0.6;}
.field input{
  width:100%;
  background:rgba(255,255,255,0.03);
  border:1px solid var(--border);
  border-radius:10px;
  padding:11px 14px;
  color:var(--text);
  font-family:'Sora',sans-serif;
  font-size:13px;
  font-weight:500;
  outline:none;
  transition:all 0.2s;
  letter-spacing:0.3px;
}
.field input::placeholder{color:var(--text3);font-weight:400;}
.field input:focus{
  border-color:rgba(99,179,237,0.4);
  background:rgba(59,130,246,0.04);
  box-shadow:0 0 0 3px rgba(59,130,246,0.08);
}

/* SUBMIT */
.submit-wrap{margin-top:6px;}
.submit-btn{
  width:100%;padding:13px;border:none;border-radius:12px;
  color:white;font-family:'Sora',sans-serif;font-size:13px;font-weight:700;
  letter-spacing:0.5px;cursor:pointer;
  position:relative;overflow:hidden;
  transition:all 0.25s cubic-bezier(0.4,0,0.2,1);
  background:linear-gradient(135deg,#f43f5e,#dc2626);
  box-shadow:0 8px 24px rgba(244,63,94,0.3),inset 0 1px 0 rgba(255,255,255,0.1);
}
.submit-btn.hosp-mode{
  background:linear-gradient(135deg,#3b82f6,#2563eb);
  box-shadow:0 8px 24px rgba(59,130,246,0.3),inset 0 1px 0 rgba(255,255,255,0.1);
}
.submit-btn::before{
  content:'';position:absolute;inset:0;
  background:linear-gradient(135deg,rgba(255,255,255,0.12),transparent);
  opacity:0;transition:0.2s;
}
.submit-btn:hover{transform:translateY(-2px);box-shadow:0 14px 32px rgba(244,63,94,0.4),inset 0 1px 0 rgba(255,255,255,0.15);}
.submit-btn.hosp-mode:hover{box-shadow:0 14px 32px rgba(59,130,246,0.4),inset 0 1px 0 rgba(255,255,255,0.15);}
.submit-btn:hover::before{opacity:1;}
.submit-btn:active{transform:translateY(0);}
.btn-inner{display:flex;align-items:center;justify-content:center;gap:8px;}
.btn-arrow{font-size:16px;transition:transform 0.2s;}
.submit-btn:hover .btn-arrow{transform:translateX(3px);}

/* STATUS BAR */
.status-bar{
  display:flex;align-items:center;justify-content:space-between;
  margin-top:20px;padding-top:16px;border-top:1px solid var(--border);
}
.status-left{display:flex;align-items:center;gap:6px;font-size:11px;color:var(--text3);}
.status-dot{width:6px;height:6px;border-radius:50%;background:#4ade80;box-shadow:0 0 8px #4ade80;animation:pulse-dot 2s infinite;}
@keyframes pulse-dot{0%,100%{box-shadow:0 0 8px #4ade80;}50%{box-shadow:0 0 14px #4ade80,0 0 24px rgba(74,222,128,0.3);}}
.status-right{font-family:'JetBrains Mono',monospace;font-size:10px;color:var(--text3);}
.ver-badge{
  display:inline-flex;align-items:center;gap:4px;
  background:rgba(255,255,255,0.04);border:1px solid var(--border);
  border-radius:6px;padding:3px 8px;
  font-family:'JetBrains Mono',monospace;font-size:10px;color:var(--text3);
}
</style>
</head>
<body>
<div class="bg-canvas">
  <div class="bg-grid"></div>
  <div class="bg-radial"></div>
  <div class="bg-scan"></div>
</div>

<!-- LEFT SIDE INFO -->
<div class="side-info left">
  <div class="info-chip">
    <div class="chip-label">System Status</div>
    <div class="chip-value green"><span class="chip-dot green"></span>All Systems Nominal</div>
  </div>
  <div class="info-chip">
    <div class="chip-label">Active Incidents</div>
    <div class="chip-value red"><span class="chip-dot red"></span>2 Active</div>
  </div>
  <div class="info-chip">
    <div class="chip-label">Response Units</div>
    <div class="chip-value amber"><span class="chip-dot amber"></span>6 / 8 Available</div>
  </div>
  <div class="info-chip">
    <div class="chip-label">Uptime</div>
    <div class="chip-value">99.97%</div>
  </div>
  <div class="info-chip">
    <div class="chip-label">Coverage Zone</div>
    <div class="chip-value">Tamil Nadu</div>
  </div>
</div>

<!-- RIGHT SIDE INFO -->
<div class="side-info right">
  <div class="info-chip">
    <div class="chip-label">Hospitals Online</div>
    <div class="chip-value green"><span class="chip-dot green"></span>5 / 5</div>
  </div>
  <div class="info-chip">
    <div class="chip-label">Police Stations</div>
    <div class="chip-value green"><span class="chip-dot green"></span>10 Active</div>
  </div>
  <div class="info-chip">
    <div class="chip-label">Today's Cases</div>
    <div class="chip-value" id="side-cases">Loading...</div>
  </div>
  <div class="info-chip">
    <div class="chip-label">Avg Response</div>
    <div class="chip-value amber">42s Avg</div>
  </div>
  <div class="info-chip">
    <div class="chip-label">Last Updated</div>
    <div class="chip-value" id="side-time">--:--:--</div>
  </div>
</div>

<!-- MAIN LOGIN -->
<div class="wrap">
  <div class="logo-area">
    <div class="logo-mark">🚨</div>
    <div class="logo-wordmark">Acci<span>Sense</span></div>
    <div class="logo-tagline">Emergency Response Intelligence</div>
  </div>

  <div class="card">
    <div class="role-label">Select Access Portal</div>
    <div class="role-grid">
      <div class="role-btn selected-police" id="btn-police" onclick="selectRole('police')">
        <div class="sel-check">✓</div>
        <span class="role-icon">👮</span>
        <div class="role-name">Police</div>
        <div class="role-desc">Command Center</div>
      </div>
      <div class="role-btn" id="btn-hospital" onclick="selectRole('hospital')">
        <div class="sel-check">✓</div>
        <span class="role-icon">🏥</span>
        <div class="role-name">Hospital</div>
        <div class="role-desc">Emergency Unit</div>
      </div>
    </div>

    <div class="divider"></div>

    <form method="post" action="/login" id="loginForm">
      <input type="hidden" name="role" id="roleInput" value="police">
      <div class="field">
        <label><span class="field-ic">🪪</span> Unit Identifier</label>
        <input type="text" name="unit_id" value="UNIT-001" placeholder="Enter unit ID" autocomplete="off">
      </div>
      <div class="field">
        <label><span class="field-ic">🔐</span> Access Code</label>
        <input type="password" name="password" value="admin123" placeholder="Enter access code">
      </div>
      <div class="submit-wrap">
        <button class="submit-btn" id="submitBtn" type="submit">
          <div class="btn-inner">
            <span id="btnLabel">Access Police Portal</span>
            <span class="btn-arrow">→</span>
          </div>
        </button>
      </div>
    </form>

    <div class="status-bar">
      <div class="status-left">
        <div class="status-dot"></div>
        <span>System Online · Real-time Active</span>
      </div>
      <div class="ver-badge">v2.4.1</div>
    </div>
  </div>
</div>

<script>
let selected = 'police';

function selectRole(role) {
  selected = role;
  document.getElementById('roleInput').value = role;

  const pb = document.getElementById('btn-police');
  const hb = document.getElementById('btn-hospital');
  const sb = document.getElementById('submitBtn');
  const lbl = document.getElementById('btnLabel');

  pb.className = 'role-btn' + (role === 'police' ? ' selected-police' : '');
  hb.className = 'role-btn' + (role === 'hospital' ? ' selected-hospital' : '');

  if (role === 'police') {
    sb.className = 'submit-btn';
    lbl.textContent = 'Access Police Portal';
  } else {
    sb.className = 'submit-btn hosp-mode';
    lbl.textContent = 'Access Hospital Portal';
  }
}

// Clock
function tick() {
  const t = new Date().toLocaleTimeString('en-IN', {hour12: false});
  const el = document.getElementById('side-time');
  if (el) el.textContent = t;
}
tick();
setInterval(tick, 1000);

// Fetch case count
fetch('/api/state').then(r => r.json()).then(s => {
  const el = document.getElementById('side-cases');
  if (el) el.textContent = s.total_cases + ' Recorded';
}).catch(() => {
  const el = document.getElementById('side-cases');
  if (el) el.textContent = '0 Recorded';
});
</script>
</body>
</html>"""

POLICE_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AcciSense — Police Dashboard</title>
<link href="https://fonts.googleapis.com/css2?family=Sora:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box;}
:root{
  --bg:#f0f4f8;
  --sidebar:#080c14;
  --sidebar2:#0d1421;
  --surface:#ffffff;
  --border:#e2e8f0;
  --accent:#f43f5e;
  --accent2:#3b82f6;
  --teal:#14b8a6;
  --warn:#f59e0b;
  --danger:#ef4444;
  --success:#22c55e;
  --text:#0f172a;
  --muted:#64748b;
  --muted2:#94a3b8;
  --sw:230px;
  --topbar-h:60px;
}
body{background:var(--bg);font-family:'Sora',sans-serif;color:var(--text);display:flex;min-height:100vh;}

/* SIDEBAR */
.sidebar{
  width:var(--sw);background:var(--sidebar);
  display:flex;flex-direction:column;
  position:fixed;top:0;left:0;height:100vh;z-index:100;
  border-right:1px solid rgba(255,255,255,0.04);
}
.sb-logo{padding:20px 18px;border-bottom:1px solid rgba(255,255,255,0.05);}
.sb-logo-row{display:flex;align-items:center;gap:10px;}
.sb-icon{
  width:36px;height:36px;
  background:linear-gradient(135deg,#f43f5e,#dc2626);
  border-radius:10px;display:flex;align-items:center;justify-content:center;
  font-size:16px;flex-shrink:0;
  box-shadow:0 4px 12px rgba(244,63,94,0.3);
}
.sb-name{font-family:'Sora',sans-serif;font-size:15px;font-weight:800;color:white;letter-spacing:-0.3px;}
.sb-sub{font-size:9px;color:#334155;letter-spacing:1.5px;margin-top:1px;text-transform:uppercase;font-weight:500;}
.sb-section{
  padding:18px 18px 6px;
  font-size:9px;font-weight:700;letter-spacing:2px;
  color:#1e293b;text-transform:uppercase;
}
.nav{
  display:flex;align-items:center;gap:9px;
  padding:9px 14px 9px 18px;
  cursor:pointer;transition:all 0.15s;
  font-size:12px;font-weight:500;
  color:#475569;margin:1px 8px;border-radius:8px;
  position:relative;
}
.nav:hover{background:rgba(255,255,255,0.04);color:#94a3b8;}
.nav.active{background:rgba(244,63,94,0.1);color:#f43f5e;}
.nav.active::before{content:'';position:absolute;left:-8px;top:50%;transform:translateY(-50%);width:3px;height:20px;background:#f43f5e;border-radius:0 3px 3px 0;}
.nav-ic{font-size:14px;width:18px;text-align:center;flex-shrink:0;}
.nav-badge{margin-left:auto;background:#ef4444;color:white;border-radius:20px;padding:1px 7px;font-size:9px;font-weight:700;animation:badgePulse 1.5s infinite;}
@keyframes badgePulse{0%,100%{box-shadow:0 0 0 0 rgba(239,68,68,0.4);}50%{box-shadow:0 0 0 4px rgba(239,68,68,0);}}
.sb-bottom{margin-top:auto;padding:14px 18px;border-top:1px solid rgba(255,255,255,0.04);}
.sb-user{display:flex;align-items:center;gap:9px;}
.av{width:34px;height:34px;background:linear-gradient(135deg,#f43f5e,#3b82f6);border-radius:9px;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:800;color:white;flex-shrink:0;}
.un{font-size:12px;font-weight:600;color:#e2e8f0;}
.ur{font-size:9px;color:#334155;letter-spacing:1px;margin-top:1px;text-transform:uppercase;}

/* TOPBAR */
.topbar{
  position:fixed;top:0;left:var(--sw);right:0;height:var(--topbar-h);
  background:white;border-bottom:1px solid var(--border);
  display:flex;align-items:center;padding:0 24px;
  z-index:50;box-shadow:0 1px 4px rgba(0,0,0,0.05);
}
.tb-breadcrumb{display:flex;align-items:center;gap:6px;}
.tb-home{font-size:11px;color:var(--muted2);}
.tb-sep{color:var(--border);font-size:14px;}
.tb-page{font-size:13px;font-weight:600;color:var(--text);}
.tb-right{margin-left:auto;display:flex;align-items:center;gap:12px;}
.live-pill{
  display:flex;align-items:center;gap:5px;
  background:#fff1f2;border:1px solid #fecdd3;
  border-radius:20px;padding:4px 10px;
  font-size:10px;font-weight:700;color:#f43f5e;letter-spacing:1px;
}
.lp-dot{width:6px;height:6px;border-radius:50%;background:#f43f5e;animation:blink 1s infinite;}
@keyframes blink{0%,100%{opacity:1;}50%{opacity:0.3;}}
.tb-time{font-family:'JetBrains Mono',monospace;font-size:12px;color:var(--muted);font-weight:500;}
.tb-logout{
  background:#f8fafc;border:1px solid var(--border);
  color:var(--muted);padding:6px 12px;border-radius:7px;
  cursor:pointer;font-size:11px;font-weight:600;
  text-decoration:none;transition:0.15s;font-family:'Sora',sans-serif;
}
.tb-logout:hover{background:#f1f5f9;color:var(--text);}

/* MAIN */
.main{margin-left:var(--sw);margin-top:var(--topbar-h);flex:1;padding:24px;overflow-y:auto;min-height:calc(100vh - var(--topbar-h));}

/* TOAST */
.toast{
  position:fixed;top:72px;right:20px;
  background:white;border:1px solid #fecdd3;border-left:3px solid #f43f5e;
  border-radius:10px;padding:12px 16px;
  font-size:12px;font-weight:600;color:var(--text);
  z-index:9999;box-shadow:0 8px 32px rgba(0,0,0,0.1);
  display:none;max-width:320px;gap:8px;align-items:center;
  animation:slideIn 0.3s ease;
}
.toast.show{display:flex;}
@keyframes slideIn{from{opacity:0;transform:translateX(16px);}to{opacity:1;transform:translateX(0);}}

/* PAGE HEADER */
.page-hdr{margin-bottom:22px;}
.page-title{font-size:20px;font-weight:800;color:var(--text);letter-spacing:-0.4px;margin-bottom:3px;}
.page-sub{font-size:12px;color:var(--muted);font-weight:400;}

/* STAT CARDS */
.stats-row{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:22px;}
.stat-card{
  background:var(--surface);border:1px solid var(--border);
  border-radius:14px;padding:18px;
  position:relative;overflow:hidden;
  transition:box-shadow 0.2s,transform 0.2s;
}
.stat-card:hover{box-shadow:0 6px 20px rgba(0,0,0,0.07);transform:translateY(-1px);}
.stat-card::after{content:'';position:absolute;top:0;left:0;right:0;height:2px;}
.stat-card.red::after{background:linear-gradient(90deg,#f43f5e,#fb7185);}
.stat-card.blue::after{background:linear-gradient(90deg,#3b82f6,#60a5fa);}
.stat-card.green::after{background:linear-gradient(90deg,#22c55e,#4ade80);}
.stat-card.amber::after{background:linear-gradient(90deg,#f59e0b,#fbbf24);}
.stat-top{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:10px;}
.stat-label{font-size:11px;font-weight:600;color:var(--muted);letter-spacing:0.2px;}
.stat-icon-box{width:36px;height:36px;border-radius:9px;display:flex;align-items:center;justify-content:center;font-size:16px;}
.stat-value{font-family:'JetBrains Mono',monospace;font-size:28px;font-weight:700;color:var(--text);line-height:1;letter-spacing:-1px;}
.stat-sub{font-size:10px;color:var(--muted2);margin-top:5px;}

/* ALERT CARD */
.alert-card{background:var(--surface);border:1px solid var(--border);border-radius:16px;padding:22px;margin-bottom:18px;}
.alert-card.emergency{border-color:#fecdd3;border-left:3px solid #f43f5e;background:#fffbfb;}
.alert-header{display:flex;align-items:center;gap:10px;margin-bottom:18px;flex-wrap:wrap;}
.emergency-pill{
  background:#fff1f2;border:1px solid #fecdd3;color:#f43f5e;
  border-radius:6px;padding:4px 10px;font-size:10px;font-weight:700;
  letter-spacing:1px;animation:flash 1.2s infinite;
}
@keyframes flash{0%,100%{opacity:1;}50%{opacity:0.4;}}
.alert-title{font-size:14px;font-weight:700;color:var(--text);}
.alert-status{margin-left:auto;font-size:10px;font-weight:700;padding:4px 10px;border-radius:20px;letter-spacing:0.5px;}
.status-pending{background:#fff1f2;color:#f43f5e;}
.status-dispatched{background:#f0fdf4;color:#16a34a;}
.status-no_hospital{background:#fffbeb;color:#d97706;}

/* INFO GRID */
.info-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:14px;}
.info-field{background:#f8fafc;border:1px solid var(--border);border-radius:9px;padding:12px;}
.ifl{font-size:9px;font-weight:700;letter-spacing:1.5px;color:var(--muted2);margin-bottom:4px;text-transform:uppercase;}
.ifv{font-size:13px;font-weight:700;color:var(--text);}
.ifv.plate{font-family:'JetBrains Mono',monospace;font-size:13px;color:#3b82f6;letter-spacing:1px;}
.ifv.police{color:#f59e0b;}

/* MAP */
.map-iframe{width:100%;height:220px;border:none;border-radius:10px;}

/* CHARTS */
.charts-row{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:18px;}
.chart-card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:18px;}
.chart-title{font-size:12px;font-weight:700;color:var(--text);margin-bottom:14px;display:flex;align-items:center;gap:7px;}
.chart-title-ic{font-size:14px;}
canvas{max-height:200px !important;}

/* TABLE */
.table-card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:18px;margin-bottom:18px;}
.table-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px;}
.table-title{font-size:12px;font-weight:700;color:var(--text);display:flex;align-items:center;gap:7px;}
table{width:100%;border-collapse:collapse;}
thead th{
  font-size:10px;font-weight:700;letter-spacing:0.5px;
  color:var(--muted2);text-align:left;
  padding:9px 12px;border-bottom:1px solid var(--border);
  background:#f8fafc;text-transform:uppercase;
}
thead th:first-child{border-radius:7px 0 0 0;}
thead th:last-child{border-radius:0 7px 0 0;}
tbody td{padding:10px 12px;font-size:12px;border-bottom:1px solid #f8fafc;color:var(--text);}
tbody tr:last-child td{border-bottom:none;}
tbody tr:hover td{background:#fafbfc;}
.badge{display:inline-flex;align-items:center;gap:3px;border-radius:20px;padding:3px 9px;font-size:10px;font-weight:700;}
.badge-success{background:#f0fdf4;color:#16a34a;border:1px solid #bbf7d0;}
.badge-danger{background:#fff1f2;color:#f43f5e;border:1px solid #fecdd3;}
.badge-warn{background:#fffbeb;color:#d97706;border:1px solid #fed7aa;}
.badge-blue{background:#eff6ff;color:#3b82f6;border:1px solid #bfdbfe;}
.plate-chip{
  font-family:'JetBrains Mono',monospace;font-size:11px;font-weight:600;
  color:#3b82f6;background:#eff6ff;padding:2px 7px;
  border-radius:5px;letter-spacing:0.5px;
}

/* VEHICLE TYPE GRID */
.vt-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:18px;}
.vt-card{
  background:var(--surface);border:1px solid var(--border);
  border-radius:12px;padding:14px;text-align:center;
  transition:all 0.2s;
}
.vt-card:hover{box-shadow:0 4px 12px rgba(0,0,0,0.06);transform:translateY(-1px);}
.vt-icon{font-size:26px;margin-bottom:7px;}
.vt-type{font-size:11px;font-weight:700;color:var(--text);}
.vt-count{font-family:'JetBrains Mono',monospace;font-size:20px;font-weight:700;color:#3b82f6;margin-top:2px;}

/* NO CASE */
.no-case{display:flex;flex-direction:column;align-items:center;justify-content:center;padding:48px;text-align:center;}
.no-case-icon{font-size:40px;margin-bottom:12px;opacity:0.2;}
.no-case-text{font-size:13px;font-weight:600;color:var(--muted);letter-spacing:0.2px;}

@keyframes pulse{0%,100%{opacity:1;transform:scale(1);}50%{opacity:0.4;transform:scale(0.8);}}
</style>
</head>
<body>
<div class="toast" id="toast"><span>🚨</span><span id="toast-msg"></span></div>

<!-- SIDEBAR -->
<div class="sidebar">
  <div class="sb-logo">
    <div class="sb-logo-row">
      <div class="sb-icon">🚨</div>
      <div>
        <div class="sb-name">AcciSense</div>
        <div class="sb-sub">Police Portal</div>
      </div>
    </div>
  </div>
  <div class="sb-section">Command</div>
  <div class="nav active" onclick="showSection('dashboard',this)"><span class="nav-ic">📊</span> Overview</div>
  <div class="nav" onclick="showSection('live',this)"><span class="nav-ic">🚨</span> Live Incident <span class="nav-badge" id="live-badge" style="display:none">!</span></div>
  <div class="nav" onclick="showSection('vehicles',this)"><span class="nav-ic">🚗</span> Vehicle Types</div>
  <div class="nav" onclick="showSection('analytics',this)"><span class="nav-ic">📈</span> Analytics</div>
  <div class="nav" onclick="showSection('history',this)"><span class="nav-ic">📋</span> Case History</div>
  <div class="sb-section">Operations</div>
  <div class="nav" onclick="showSection('fleet',this)"><span class="nav-ic">🚓</span> Fleet Status</div>
  <div class="nav" onclick="showSection('zones',this)"><span class="nav-ic">🗺️</span> Zone Map</div>
  <div class="nav" onclick="showSection('cctv',this)"><span class="nav-ic">📡</span> CCTV Feed</div>
  <div class="sb-bottom">
    <div class="sb-user">
      <div class="av">OP</div>
      <div>
        <div class="un">Officer Admin</div>
        <div class="ur">Unit-001 · Police</div>
      </div>
    </div>
  </div>
</div>

<!-- TOPBAR -->
<div class="topbar">
  <div class="tb-breadcrumb">
    <span class="tb-home">AcciSense</span>
    <span class="tb-sep">/</span>
    <span class="tb-page" id="tb-page-name">Overview</span>
  </div>
  <div class="tb-right">
    <div class="live-pill"><div class="lp-dot"></div>LIVE</div>
    <div class="tb-time" id="clock">--:--:--</div>
    <a href="/" class="tb-logout">← Logout</a>
  </div>
</div>

<!-- MAIN -->
<div class="main">

  <!-- DASHBOARD -->
  <div id="section-dashboard">
    <div class="page-hdr">
      <div class="page-title">Command Overview</div>
      <div class="page-sub">Real-time incident monitoring and response coordination</div>
    </div>
    <div class="stats-row">
      <div class="stat-card blue">
        <div class="stat-top"><div class="stat-label">Total Cases Today</div><div class="stat-icon-box" style="background:#eff6ff">📊</div></div>
        <div class="stat-value" id="st-total">0</div>
        <div class="stat-sub">All incidents logged</div>
      </div>
      <div class="stat-card red">
        <div class="stat-top"><div class="stat-label">Active Incidents</div><div class="stat-icon-box" style="background:#fff1f2">🚨</div></div>
        <div class="stat-value" id="st-active" style="color:#f43f5e">0</div>
        <div class="stat-sub">Requiring response</div>
      </div>
      <div class="stat-card green">
        <div class="stat-top"><div class="stat-label">Resolved Today</div><div class="stat-icon-box" style="background:#f0fdf4">✅</div></div>
        <div class="stat-value" id="st-resolved" style="color:#22c55e">0</div>
        <div class="stat-sub">Ambulance dispatched</div>
      </div>
      <div class="stat-card amber">
        <div class="stat-top"><div class="stat-label">Avg Response</div><div class="stat-icon-box" style="background:#fffbeb">⏱️</div></div>
        <div class="stat-value" id="st-avgtime" style="color:#f59e0b">--</div>
        <div class="stat-sub">Seconds avg</div>
      </div>
    </div>
    <div id="dash-alert"></div>
    <div class="charts-row">
      <div class="chart-card"><div class="chart-title"><span class="chart-title-ic">📈</span> Incident Trend</div><canvas id="trendChart"></canvas></div>
      <div class="chart-card"><div class="chart-title"><span class="chart-title-ic">🚗</span> Vehicle Types Involved</div><canvas id="vtChart"></canvas></div>
    </div>
    <div class="table-card">
      <div class="table-header"><div class="table-title"><span>📋</span> Recent Incidents</div></div>
      <table>
        <thead><tr><th>Time</th><th>Vehicle 1</th><th>Vehicle 2</th><th>Type</th><th>Police Station</th><th>Hospital</th><th>Status</th></tr></thead>
        <tbody id="dash-tbody"><tr><td colspan="7" style="text-align:center;color:var(--muted2);padding:28px;font-size:12px;">No incidents recorded yet</td></tr></tbody>
      </table>
    </div>
  </div>

  <!-- LIVE -->
  <div id="section-live" style="display:none">
    <div class="page-hdr"><div class="page-title">Live Incident</div><div class="page-sub">Real-time accident monitoring and response</div></div>
    <div id="live-detail"></div>
  </div>

  <!-- VEHICLE TYPES -->
  <div id="section-vehicles" style="display:none">
    <div class="page-hdr"><div class="page-title">Vehicle Analysis</div><div class="page-sub">Breakdown of vehicle types involved in accidents</div></div>
    <div class="vt-grid" id="vt-grid"></div>
    <div class="charts-row">
      <div class="chart-card"><div class="chart-title"><span class="chart-title-ic">🍩</span> Type Distribution</div><canvas id="vtDoughnut"></canvas></div>
      <div class="chart-card"><div class="chart-title"><span class="chart-title-ic">📊</span> Incidents by Type</div><canvas id="vtBar"></canvas></div>
    </div>
    <div class="table-card">
      <div class="table-header"><div class="table-title"><span>🚗</span> Vehicle Incident Log</div></div>
      <table>
        <thead><tr><th>Time</th><th>Plate</th><th>Type</th><th>Icon</th><th>Hospital</th><th>Status</th></tr></thead>
        <tbody id="vt-tbody"><tr><td colspan="6" style="text-align:center;color:var(--muted2);padding:28px;font-size:12px;">No data</td></tr></tbody>
      </table>
    </div>
  </div>

  <!-- ANALYTICS -->
  <div id="section-analytics" style="display:none">
    <div class="page-hdr"><div class="page-title">Analytics</div><div class="page-sub">Incident patterns and performance metrics</div></div>
    <div class="charts-row">
      <div class="chart-card"><div class="chart-title"><span class="chart-title-ic">🚓</span> Cases by Police Station</div><canvas id="stationChart"></canvas></div>
      <div class="chart-card"><div class="chart-title"><span class="chart-title-ic">⏱️</span> Response Time Distribution</div><canvas id="responseChart"></canvas></div>
    </div>
    <div class="charts-row">
      <div class="chart-card"><div class="chart-title"><span class="chart-title-ic">📅</span> Weekly Incidents</div><canvas id="weeklyChart"></canvas></div>
      <div class="chart-card"><div class="chart-title"><span class="chart-title-ic">🕐</span> Incidents by Hour</div><canvas id="hourChart"></canvas></div>
    </div>
  </div>

  <!-- HISTORY -->
  <div id="section-history" style="display:none">
    <div class="page-hdr"><div class="page-title">Case History</div><div class="page-sub">Complete log of all reported incidents</div></div>
    <div class="table-card">
      <table>
        <thead><tr><th>Date</th><th>Time</th><th>Vehicle 1</th><th>Vehicle 2</th><th>Types</th><th>Police Station</th><th>Hospital</th><th>Response</th><th>Status</th></tr></thead>
        <tbody id="hist-tbody"><tr><td colspan="9" style="text-align:center;color:var(--muted2);padding:28px;font-size:12px;">No cases yet</td></tr></tbody>
      </table>
    </div>
  </div>

  <!-- FLEET STATUS -->
  <div id="section-fleet" style="display:none">
    <div class="page-hdr"><div class="page-title">Fleet Status</div><div class="page-sub">Live status of all emergency and patrol vehicles</div></div>
    <div class="stats-row">
      <div class="stat-card blue"><div class="stat-top"><div class="stat-label">Total Vehicles</div><div class="stat-icon-box" style="background:#eff6ff">🚗</div></div><div class="stat-value" id="fleet-total">8</div><div class="stat-sub">In fleet</div></div>
      <div class="stat-card green"><div class="stat-top"><div class="stat-label">Available</div><div class="stat-icon-box" style="background:#f0fdf4">✅</div></div><div class="stat-value" id="fleet-avail" style="color:#22c55e">0</div><div class="stat-sub">Ready to dispatch</div></div>
      <div class="stat-card blue"><div class="stat-top"><div class="stat-label">Dispatched</div><div class="stat-icon-box" style="background:#eff6ff">🚑</div></div><div class="stat-value" id="fleet-disp" style="color:#3b82f6">0</div><div class="stat-sub">On active duty</div></div>
      <div class="stat-card amber"><div class="stat-top"><div class="stat-label">Maintenance</div><div class="stat-icon-box" style="background:#fffbeb">🔧</div></div><div class="stat-value" id="fleet-maint" style="color:#f59e0b">0</div><div class="stat-sub">Under service</div></div>
    </div>
    <div class="table-card">
      <div class="table-header"><div class="table-title"><span>🚓</span> Vehicle Fleet</div></div>
      <table>
        <thead><tr><th>Vehicle ID</th><th>Type</th><th>Driver</th><th>Current Location</th><th>Last Trip</th><th>Status</th></tr></thead>
        <tbody id="fleet-tbody"></tbody>
      </table>
    </div>
  </div>

  <!-- ZONE MAP -->
  <div id="section-zones" style="display:none">
    <div class="page-hdr"><div class="page-title">Zone Map</div><div class="page-sub">Geographic coverage and incident mapping</div></div>
    <div class="alert-card">
      <iframe src="https://maps.google.com/maps?q=13.0827,80.2707&z=13&output=embed" style="width:100%;height:460px;border:none;border-radius:10px;"></iframe>
    </div>
  </div>

  <!-- CCTV -->
  <div id="section-cctv" style="display:none">
    <div class="page-hdr"><div class="page-title">CCTV Feed</div><div class="page-sub">Surveillance camera monitoring across city zones</div></div>
    <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:14px;">
      <div class="chart-card" style="text-align:center;padding:36px 18px;"><div style="font-size:34px;margin-bottom:10px;opacity:0.25;">📷</div><div style="font-size:12px;font-weight:700;color:var(--muted);">Camera 01 — NH-44 Junction</div><div style="font-size:10px;color:#22c55e;margin-top:6px;font-weight:600;">● Online</div></div>
      <div class="chart-card" style="text-align:center;padding:36px 18px;"><div style="font-size:34px;margin-bottom:10px;opacity:0.25;">📷</div><div style="font-size:12px;font-weight:700;color:var(--muted);">Camera 02 — City Mall Signal</div><div style="font-size:10px;color:#22c55e;margin-top:6px;font-weight:600;">● Online</div></div>
      <div class="chart-card" style="text-align:center;padding:36px 18px;"><div style="font-size:34px;margin-bottom:10px;opacity:0.25;">📷</div><div style="font-size:12px;font-weight:700;color:var(--muted);">Camera 03 — Railway Gate</div><div style="font-size:10px;color:#f43f5e;margin-top:6px;font-weight:600;">● Offline</div></div>
      <div class="chart-card" style="text-align:center;padding:36px 18px;"><div style="font-size:34px;margin-bottom:10px;opacity:0.25;">📷</div><div style="font-size:12px;font-weight:700;color:var(--muted);">Camera 04 — Bus Terminal</div><div style="font-size:10px;color:#22c55e;margin-top:6px;font-weight:600;">● Online</div></div>
      <div class="chart-card" style="text-align:center;padding:36px 18px;"><div style="font-size:34px;margin-bottom:10px;opacity:0.25;">📷</div><div style="font-size:12px;font-weight:700;color:var(--muted);">Camera 05 — Airport Road</div><div style="font-size:10px;color:#22c55e;margin-top:6px;font-weight:600;">● Online</div></div>
      <div class="chart-card" style="text-align:center;padding:36px 18px;"><div style="font-size:34px;margin-bottom:10px;opacity:0.25;">📷</div><div style="font-size:12px;font-weight:700;color:var(--muted);">Camera 06 — Industrial Zone</div><div style="font-size:10px;color:#f59e0b;margin-top:6px;font-weight:600;">● Degraded</div></div>
    </div>
  </div>

</div>

<script>
let CS={}, lastCases=0;
let trendC, vtC, stationC, responseC, weeklyC, hourC, vtDC, vtBC;
let vtCounts={}, stationCounts={}, hourCounts=new Array(24).fill(0), responseTimes=[];

setInterval(()=>document.getElementById('clock').textContent=new Date().toLocaleTimeString('en-IN',{hour12:false}),1000);

function showSection(name,el){
  document.querySelectorAll('[id^="section-"]').forEach(s=>s.style.display='none');
  document.getElementById('section-'+name).style.display='block';
  document.querySelectorAll('.nav').forEach(n=>n.classList.remove('active'));
  if(el) el.classList.add('active');
  const names={dashboard:'Overview',live:'Live Incident',vehicles:'Vehicle Types',analytics:'Analytics',history:'Case History',fleet:'Fleet Status',zones:'Zone Map',cctv:'CCTV Feed'};
  document.getElementById('tb-page-name').textContent=names[name]||name;
  if(name==='analytics') initAnalyticsCharts();
  if(name==='history') renderHistory();
  if(name==='vehicles') renderVehicles();
  if(name==='fleet') renderFleet();
}

function toast(msg){const t=document.getElementById('toast');document.getElementById('toast-msg').textContent=msg;t.classList.add('show');setTimeout(()=>t.classList.remove('show'),6000);}

const VT={TN01:{type:"Sedan",icon:"🚗"},TN02:{type:"SUV",icon:"🚙"},TN03:{type:"Truck",icon:"🚛"},MH:{type:"Motorcycle",icon:"🏍️"},KA:{type:"Auto-Rickshaw",icon:"🛺"},AP:{type:"Bus",icon:"🚌"},TS:{type:"Van",icon:"🚐"},DL:{type:"Hatchback",icon:"🚘"},GJ:{type:"Pickup",icon:"🛻"},RJ:{type:"Minibus",icon:"🚎"},HR:{type:"Bicycle",icon:"🚲"},UP:{type:"Ambulance",icon:"🚑"},WB:{type:"Taxi",icon:"🚕"},MP:{type:"Police Car",icon:"🚓"},PB:{type:"Fire Truck",icon:"🚒"}};
function getVI(p){for(let k in VT){if(p.toUpperCase().startsWith(k))return VT[k];}return{type:"Vehicle",icon:"🚗"};}
function fmtTime(iso){return iso?new Date(iso).toLocaleTimeString('en-IN',{hour12:false}):'--';}

function buildAlert(s){
  if(!s.current_case)return'<div class="no-case"><div class="no-case-icon">🛡️</div><div class="no-case-text">No active incident — System monitoring</div></div>';
  const c=s.current_case;
  const v1=getVI(c.vehicle_1),v2=getVI(c.vehicle_2);
  const stClass=s.case_status==='pending'?'pending':s.case_status==='ambulance_dispatched'?'dispatched':'no_hospital';
  const mapUrl=`https://maps.google.com/maps?q=${c.latitude},${c.longitude}&z=16&output=embed`;
  return`<div class="alert-card ${s.case_status==='pending'?'emergency':''}">
    <div class="alert-header">
      ${s.case_status==='pending'?'<div class="emergency-pill">🚨 EMERGENCY ALERT</div>':'<div class="emergency-pill" style="background:#f0fdf4;border-color:#bbf7d0;color:#16a34a;animation:none">✅ RESOLVED</div>'}
      <div class="alert-title">Accident Detected — Immediate Response Required</div>
      <div class="alert-status status-${stClass}">${s.case_status.replace(/_/g,' ').toUpperCase()}</div>
    </div>
    <div class="info-grid">
      <div class="info-field"><div class="ifl">Vehicle 1 Plate</div><div class="ifv plate">${c.vehicle_1}</div></div>
      <div class="info-field"><div class="ifl">Vehicle 2 Plate</div><div class="ifv plate">${c.vehicle_2}</div></div>
      <div class="info-field"><div class="ifl">Vehicle Types</div><div class="ifv">${v1.icon} ${v1.type} / ${v2.icon} ${v2.type}</div></div>
      <div class="info-field"><div class="ifl">Police Station</div><div class="ifv police">🚓 ${s.selected_police}</div></div>
      <div class="info-field"><div class="ifl">Alert Time</div><div class="ifv">${fmtTime(s.alert_time)}</div></div>
      <div class="info-field"><div class="ifl">GPS Coordinates</div><div class="ifv" style="font-family:'JetBrains Mono',monospace;font-size:11px;">${c.latitude}, ${c.longitude}</div></div>
    </div>
    <iframe class="map-iframe" src="${mapUrl}" allowfullscreen loading="lazy"></iframe>
  </div>`;
}

function initCharts(){
  Chart.defaults.font.family='Sora';Chart.defaults.color='#64748b';Chart.defaults.font.size=11;
  trendC=new Chart(document.getElementById('trendChart'),{type:'line',data:{labels:[],datasets:[{label:'Incidents',data:[],borderColor:'#f43f5e',backgroundColor:'rgba(244,63,94,0.06)',tension:0.4,fill:true,pointBackgroundColor:'#f43f5e',pointRadius:3,borderWidth:2}]},options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,grid:{color:'rgba(0,0,0,0.04)'}},x:{grid:{display:false}}},animation:false}});
  vtC=new Chart(document.getElementById('vtChart'),{type:'doughnut',data:{labels:[],datasets:[{data:[],backgroundColor:['#f43f5e','#3b82f6','#14b8a6','#f59e0b','#8b5cf6','#22c55e','#fb7185','#60a5fa'],borderWidth:0}]},options:{plugins:{legend:{position:'bottom',labels:{boxWidth:8,padding:8,font:{size:10}}}},cutout:'62%',animation:false}});
}

function initAnalyticsCharts(){
  if(stationC)return;
  const labels=Object.keys(stationCounts).length?Object.keys(stationCounts):['No Data'];
  const vals=Object.values(stationCounts).length?Object.values(stationCounts):[0];
  stationC=new Chart(document.getElementById('stationChart'),{type:'bar',data:{labels,datasets:[{label:'Cases',data:vals,backgroundColor:'rgba(244,63,94,0.5)',borderColor:'#f43f5e',borderWidth:1,borderRadius:5}]},options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,grid:{color:'rgba(0,0,0,0.04)'}},x:{grid:{display:false},ticks:{maxRotation:45,font:{size:9}}}},animation:false}});
  responseC=new Chart(document.getElementById('responseChart'),{type:'bar',data:{labels:['0-30s','31-60s','61-120s','120s+'],datasets:[{label:'Cases',data:[0,0,0,0],backgroundColor:['rgba(34,197,94,0.5)','rgba(59,130,246,0.5)','rgba(245,158,11,0.5)','rgba(244,63,94,0.5)'],borderWidth:0,borderRadius:5}]},options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,grid:{color:'rgba(0,0,0,0.04)'}},x:{grid:{display:false}}},animation:false}});
  weeklyC=new Chart(document.getElementById('weeklyChart'),{type:'line',data:{labels:['Mon','Tue','Wed','Thu','Fri','Sat','Sun'],datasets:[{label:'Incidents',data:[0,0,0,0,0,0,0],borderColor:'#3b82f6',backgroundColor:'rgba(59,130,246,0.05)',tension:0.4,fill:true,borderWidth:2}]},options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,grid:{color:'rgba(0,0,0,0.04)'}},x:{grid:{display:false}}},animation:false}});
  hourC=new Chart(document.getElementById('hourChart'),{type:'bar',data:{labels:Array.from({length:24},(_,i)=>i+':00'),datasets:[{label:'Incidents',data:hourCounts,backgroundColor:'rgba(139,92,246,0.5)',borderColor:'#8b5cf6',borderWidth:1,borderRadius:3}]},options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,grid:{color:'rgba(0,0,0,0.04)'}},x:{grid:{display:false},ticks:{maxRotation:90,font:{size:8}}}},animation:false}});
}

function renderVehicles(){
  const grid=document.getElementById('vt-grid');
  if(!Object.keys(vtCounts).length){grid.innerHTML='<div style="grid-column:1/-1;text-align:center;color:var(--muted2);padding:40px;font-size:13px;">No vehicle data yet</div>';return;}
  grid.innerHTML=Object.entries(vtCounts).map(([type,count])=>{const icon=Object.values(VT).find(v=>v.type===type)?.icon||'🚗';return`<div class="vt-card"><div class="vt-icon">${icon}</div><div class="vt-type">${type}</div><div class="vt-count">${count}</div></div>`;}).join('');
  const keys=Object.keys(vtCounts),vals=Object.values(vtCounts);
  if(!vtDC){
    vtDC=new Chart(document.getElementById('vtDoughnut'),{type:'doughnut',data:{labels:keys,datasets:[{data:vals,backgroundColor:['#f43f5e','#3b82f6','#14b8a6','#f59e0b','#8b5cf6','#22c55e','#fb7185','#60a5fa'],borderWidth:0}]},options:{plugins:{legend:{position:'bottom',labels:{boxWidth:8,padding:8,font:{size:10}}}},cutout:'62%',animation:false}});
    vtBC=new Chart(document.getElementById('vtBar'),{type:'bar',data:{labels:keys,datasets:[{label:'Count',data:vals,backgroundColor:'rgba(59,130,246,0.5)',borderColor:'#3b82f6',borderWidth:1,borderRadius:5}]},options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,grid:{color:'rgba(0,0,0,0.04)'}},x:{grid:{display:false}}},animation:false}});
  }else{vtDC.data.labels=keys;vtDC.data.datasets[0].data=vals;vtDC.update();vtBC.data.labels=keys;vtBC.data.datasets[0].data=vals;vtBC.update();}
  const tbody=document.getElementById('vt-tbody');
  if(CS.case_history?.length){tbody.innerHTML=CS.case_history.flatMap(c=>[{plate:c.v1,type:c.v1type||getVI(c.v1).type,icon:c.v1icon||getVI(c.v1).icon,time:c.time,hospital:c.hospital,status:c.status},{plate:c.v2,type:c.v2type||getVI(c.v2).type,icon:c.v2icon||getVI(c.v2).icon,time:c.time,hospital:c.hospital,status:c.status}]).map(r=>`<tr><td>${r.time}</td><td><span class="plate-chip">${r.plate}</span></td><td>${r.type}</td><td style="font-size:18px">${r.icon}</td><td>${r.hospital}</td><td><span class="badge badge-success">${r.status}</span></td></tr>`).join('');}
  else{tbody.innerHTML='<tr><td colspan="6" style="text-align:center;color:var(--muted2);padding:28px;font-size:12px;">No data</td></tr>';}
}

function renderHistory(){
  const tbody=document.getElementById('hist-tbody');
  if(!CS.case_history?.length){tbody.innerHTML='<tr><td colspan="9" style="text-align:center;color:var(--muted2);padding:28px;font-size:12px;">No cases yet</td></tr>';return;}
  tbody.innerHTML=CS.case_history.slice().reverse().map(c=>`<tr><td>${c.date}</td><td>${c.time}</td><td><span class="plate-chip">${c.v1}</span></td><td><span class="plate-chip">${c.v2}</span></td><td>${(c.v1icon||getVI(c.v1).icon)} ${(c.v1type||getVI(c.v1).type)} / ${(c.v2icon||getVI(c.v2).icon)} ${(c.v2type||getVI(c.v2).type)}</td><td style="color:#f59e0b;font-size:11px;">${c.police}</td><td>${c.hospital}</td><td><span style="font-family:'JetBrains Mono',monospace;font-size:11px;">${c.response_time}s</span></td><td><span class="badge badge-success">${c.status}</span></td></tr>`).join('');
}

async function renderFleet(){
  const r=await fetch('/api/fleet-status');const data=await r.json();
  let avail=0,disp=0,maint=0;
  data.forEach(f=>{if(f.status==='available')avail++;else if(f.status==='dispatched')disp++;else maint++;});
  document.getElementById('fleet-avail').textContent=avail;
  document.getElementById('fleet-disp').textContent=disp;
  document.getElementById('fleet-maint').textContent=maint;
  document.getElementById('fleet-total').textContent=data.length;
  document.getElementById('fleet-tbody').innerHTML=data.map(f=>{
    const bc=f.status==='available'?'badge-success':f.status==='dispatched'?'badge-blue':'badge-warn';
    const lbl=f.status==='available'?'✅ Available':f.status==='dispatched'?'🚑 Dispatched':'🔧 Maintenance';
    return`<tr><td><strong style="font-family:'JetBrains Mono',monospace;font-size:11px;">${f.id}</strong></td><td>${f.type}</td><td>${f.driver}</td><td>${f.location}</td><td>${f.last_trip}</td><td><span class="badge ${bc}">${lbl}</span></td></tr>`;
  }).join('');
}

async function poll(){
  try{
    const r=await fetch('/api/state');const s=await r.json();CS=s;
    const hist=s.case_history||[];
    document.getElementById('st-total').textContent=s.total_cases;
    document.getElementById('st-active').textContent=s.case_status==='pending'?1:0;
    document.getElementById('st-resolved').textContent=hist.length;
    const avg=hist.length?Math.round(hist.reduce((a,c)=>a+c.response_time,0)/hist.length):'--';
    document.getElementById('st-avgtime').textContent=avg;
    const b=document.getElementById('live-badge');
    b.style.display=s.case_status==='pending'?'inline':'none';
    vtCounts={};stationCounts={};
    hist.forEach(c=>{[c.v1,c.v2].forEach(p=>{const vt=getVI(p).type;vtCounts[vt]=(vtCounts[vt]||0)+1;});stationCounts[c.police]=(stationCounts[c.police]||0)+1;const h=parseInt(c.time.split(':')[0]);if(!isNaN(h))hourCounts[h]++;responseTimes.push(c.response_time);});
    if(s.current_case&&s.total_cases>lastCases&&s.case_status==='pending'){const v1=getVI(s.current_case.vehicle_1),v2=getVI(s.current_case.vehicle_2);toast(`Accident: ${v1.icon}${s.current_case.vehicle_1} & ${v2.icon}${s.current_case.vehicle_2}`);lastCases=s.total_cases;}
    if(s.total_cases>0)lastCases=s.total_cases;
    document.getElementById('dash-alert').innerHTML=buildAlert(s);
    document.getElementById('live-detail').innerHTML=buildAlert(s);
    const tbody=document.getElementById('dash-tbody');
    if(hist.length){tbody.innerHTML=hist.slice(-5).reverse().map(c=>`<tr><td>${c.time}</td><td><span class="plate-chip">${c.v1}</span></td><td><span class="plate-chip">${c.v2}</span></td><td>${(c.v1icon||getVI(c.v1).icon)} ${(c.v1type||getVI(c.v1).type)}</td><td style="font-size:11px;color:#f59e0b;">${c.police}</td><td>${c.hospital}</td><td><span class="badge badge-success">${c.status}</span></td></tr>`).join('');}
    const tLabels=hist.map(c=>c.time.substr(0,5));
    trendC.data.labels=tLabels.length?tLabels:['--'];
    trendC.data.datasets[0].data=tLabels.map((_,i)=>i+1);
    trendC.update();
    const vl=Object.keys(vtCounts),vv=Object.values(vtCounts);
    vtC.data.labels=vl.length?vl:['No Data'];vtC.data.datasets[0].data=vv.length?vv:[1];
    vtC.update();
    if(document.getElementById('section-analytics').style.display!=='none'&&stationC){
      stationC.data.labels=Object.keys(stationCounts)||['No Data'];
      stationC.data.datasets[0].data=Object.values(stationCounts)||[0];
      stationC.update();
      const bins=[0,0,0,0];
      responseTimes.forEach(t=>{if(t<=30)bins[0]++;else if(t<=60)bins[1]++;else if(t<=120)bins[2]++;else bins[3]++;});
      responseC.data.datasets[0].data=bins;responseC.update();
      hourC.data.datasets[0].data=hourCounts;hourC.update();
    }
  }catch(e){}
}

initCharts();poll();setInterval(poll,2000);
</script>
</body>
</html>"""

HOSPITAL_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AcciSense — Hospital Dashboard</title>
<link href="https://fonts.googleapis.com/css2?family=Sora:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box;}
:root{
  --bg:#f0f4f8;
  --sidebar:#07110d;
  --surface:#ffffff;
  --border:#e2e8f0;
  --accent:#10b981;
  --accent2:#3b82f6;
  --warn:#f59e0b;
  --danger:#ef4444;
  --success:#22c55e;
  --text:#0f172a;
  --muted:#64748b;
  --muted2:#94a3b8;
  --sw:230px;
  --topbar-h:60px;
}
body{background:var(--bg);font-family:'Sora',sans-serif;color:var(--text);display:flex;min-height:100vh;}
.sidebar{width:var(--sw);background:var(--sidebar);display:flex;flex-direction:column;position:fixed;top:0;left:0;height:100vh;z-index:100;border-right:1px solid rgba(255,255,255,0.04);}
.sb-logo{padding:20px 18px;border-bottom:1px solid rgba(255,255,255,0.05);}
.sb-logo-row{display:flex;align-items:center;gap:10px;}
.sb-icon{width:36px;height:36px;background:linear-gradient(135deg,#10b981,#059669);border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0;box-shadow:0 4px 12px rgba(16,185,129,0.3);}
.sb-name{font-family:'Sora',sans-serif;font-size:15px;font-weight:800;color:white;letter-spacing:-0.3px;}
.sb-sub{font-size:9px;color:#1a3327;letter-spacing:1.5px;margin-top:1px;text-transform:uppercase;font-weight:500;}
.sb-section{padding:18px 18px 6px;font-size:9px;font-weight:700;letter-spacing:2px;color:#0d2619;text-transform:uppercase;}
.nav{display:flex;align-items:center;gap:9px;padding:9px 14px 9px 18px;cursor:pointer;transition:all 0.15s;font-size:12px;font-weight:500;color:#2d6a4f;margin:1px 8px;border-radius:8px;position:relative;}
.nav:hover{background:rgba(16,185,129,0.08);color:#6ee7b7;}
.nav.active{background:rgba(16,185,129,0.1);color:#10b981;}
.nav.active::before{content:'';position:absolute;left:-8px;top:50%;transform:translateY(-50%);width:3px;height:20px;background:#10b981;border-radius:0 3px 3px 0;}
.nav-ic{font-size:14px;width:18px;text-align:center;flex-shrink:0;}
.nav-badge{margin-left:auto;background:#ef4444;color:white;border-radius:20px;padding:1px 7px;font-size:9px;font-weight:700;animation:badgePulse 1.5s infinite;}
@keyframes badgePulse{0%,100%{box-shadow:0 0 0 0 rgba(239,68,68,0.4);}50%{box-shadow:0 0 0 4px rgba(239,68,68,0);}}
.sb-bottom{margin-top:auto;padding:14px 18px;border-top:1px solid rgba(255,255,255,0.04);}
.sb-user{display:flex;align-items:center;gap:9px;}
.av{width:34px;height:34px;background:linear-gradient(135deg,#10b981,#3b82f6);border-radius:9px;display:flex;align-items:center;justify-content:center;font-size:16px;color:white;flex-shrink:0;}
.un{font-size:12px;font-weight:600;color:#e2e8f0;}
.ur{font-size:9px;color:#1a3327;letter-spacing:1px;margin-top:1px;text-transform:uppercase;}
.topbar{position:fixed;top:0;left:var(--sw);right:0;height:var(--topbar-h);background:white;border-bottom:1px solid var(--border);display:flex;align-items:center;padding:0 24px;z-index:50;box-shadow:0 1px 4px rgba(0,0,0,0.05);}
.tb-breadcrumb{display:flex;align-items:center;gap:6px;}
.tb-home{font-size:11px;color:var(--muted2);}
.tb-sep{color:var(--border);font-size:14px;}
.tb-page{font-size:13px;font-weight:600;color:var(--text);}
.tb-right{margin-left:auto;display:flex;align-items:center;gap:12px;}
.live-pill{display:flex;align-items:center;gap:5px;background:#ecfdf5;border:1px solid #a7f3d0;border-radius:20px;padding:4px 10px;font-size:10px;font-weight:700;color:#059669;letter-spacing:1px;}
.lp-dot{width:6px;height:6px;border-radius:50%;background:#10b981;animation:pulse-g 2s infinite;}
@keyframes pulse-g{0%,100%{box-shadow:0 0 0 0 rgba(16,185,129,0.4);}50%{box-shadow:0 0 0 4px rgba(16,185,129,0);}}
.tb-time{font-family:'JetBrains Mono',monospace;font-size:12px;color:var(--muted);}
.tb-logout{background:#f8fafc;border:1px solid var(--border);color:var(--muted);padding:6px 12px;border-radius:7px;cursor:pointer;font-size:11px;font-weight:600;text-decoration:none;transition:0.15s;font-family:'Sora',sans-serif;}
.tb-logout:hover{background:#f1f5f9;color:var(--text);}
.main{margin-left:var(--sw);margin-top:var(--topbar-h);flex:1;padding:24px;overflow-y:auto;}
.toast{position:fixed;top:72px;right:20px;background:white;border:1px solid #a7f3d0;border-left:3px solid #10b981;border-radius:10px;padding:12px 16px;font-size:12px;font-weight:600;color:var(--text);z-index:9999;box-shadow:0 8px 32px rgba(0,0,0,0.1);display:none;max-width:320px;gap:8px;align-items:center;}
.toast.show{display:flex;animation:slideIn 0.3s ease;}
@keyframes slideIn{from{opacity:0;transform:translateX(16px);}to{opacity:1;transform:translateX(0);}}
.page-hdr{margin-bottom:22px;}
.page-title{font-size:20px;font-weight:800;color:var(--text);letter-spacing:-0.4px;margin-bottom:3px;}
.page-sub{font-size:12px;color:var(--muted);}
.stats-row{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:22px;}
.stat-card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:18px;transition:box-shadow 0.2s,transform 0.2s;position:relative;overflow:hidden;}
.stat-card:hover{box-shadow:0 6px 20px rgba(0,0,0,0.07);transform:translateY(-1px);}
.stat-card::after{content:'';position:absolute;top:0;left:0;right:0;height:2px;}
.stat-card.green::after{background:linear-gradient(90deg,#10b981,#34d399);}
.stat-card.red::after{background:linear-gradient(90deg,#ef4444,#fb7185);}
.stat-card.amber::after{background:linear-gradient(90deg,#f59e0b,#fbbf24);}
.stat-card.blue::after{background:linear-gradient(90deg,#3b82f6,#60a5fa);}
.stat-card.purple::after{background:linear-gradient(90deg,#8b5cf6,#a78bfa);}
.stat-top{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:10px;}
.stat-label{font-size:11px;font-weight:600;color:var(--muted);}
.stat-icon-box{width:36px;height:36px;border-radius:9px;display:flex;align-items:center;justify-content:center;font-size:16px;}
.stat-value{font-family:'JetBrains Mono',monospace;font-size:28px;font-weight:700;color:var(--text);line-height:1;letter-spacing:-1px;}
.stat-sub{font-size:10px;color:var(--muted2);margin-top:5px;}
.alert-card{background:var(--surface);border:1px solid var(--border);border-radius:16px;padding:22px;margin-bottom:18px;}
.alert-card.emergency{border-color:#fecdd3;border-left:3px solid #ef4444;background:#fffbfb;}
.alert-header{display:flex;align-items:center;gap:10px;margin-bottom:18px;flex-wrap:wrap;}
.emergency-pill{background:#fff1f2;border:1px solid #fecdd3;color:#ef4444;border-radius:6px;padding:4px 10px;font-size:10px;font-weight:700;letter-spacing:1px;animation:flash 1.2s infinite;}
@keyframes flash{0%,100%{opacity:1;}50%{opacity:0.4;}}
.alert-title{font-size:14px;font-weight:700;color:var(--text);}
.alert-status{margin-left:auto;font-size:10px;font-weight:700;padding:4px 10px;border-radius:20px;letter-spacing:0.5px;}
.status-pending{background:#fff1f2;color:#ef4444;}
.status-dispatched{background:#ecfdf5;color:#059669;}
.status-no_hospital{background:#fffbeb;color:#d97706;}
.info-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:14px;}
.info-field{background:#f8fafc;border:1px solid var(--border);border-radius:9px;padding:12px;}
.ifl{font-size:9px;font-weight:700;letter-spacing:1.5px;color:var(--muted2);margin-bottom:4px;text-transform:uppercase;}
.ifv{font-size:13px;font-weight:700;color:var(--text);}
.action-btns{display:flex;gap:12px;margin-top:4px;}
.btn-accept{flex:1;padding:13px;background:linear-gradient(135deg,#10b981,#059669);border:none;border-radius:11px;color:white;font-family:'Sora',sans-serif;font-size:13px;font-weight:700;cursor:pointer;transition:all 0.2s;box-shadow:0 6px 18px rgba(16,185,129,0.3);letter-spacing:0.3px;}
.btn-accept:hover{transform:translateY(-2px);box-shadow:0 10px 26px rgba(16,185,129,0.4);}
.btn-decline{flex:1;padding:13px;background:linear-gradient(135deg,#ef4444,#dc2626);border:none;border-radius:11px;color:white;font-family:'Sora',sans-serif;font-size:13px;font-weight:700;cursor:pointer;transition:all 0.2s;box-shadow:0 6px 18px rgba(239,68,68,0.3);letter-spacing:0.3px;}
.btn-decline:hover{transform:translateY(-2px);box-shadow:0 10px 26px rgba(239,68,68,0.4);}
.charts-row{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:18px;}
.chart-card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:18px;}
.chart-title{font-size:12px;font-weight:700;color:var(--text);margin-bottom:14px;display:flex;align-items:center;gap:7px;}
canvas{max-height:220px !important;}
.map-iframe{width:100%;height:220px;border:none;border-radius:10px;}
.bed-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:18px;}
.bed-card{background:var(--surface);border:1px solid var(--border);border-radius:13px;padding:16px;position:relative;overflow:hidden;}
.bed-card::after{content:'';position:absolute;bottom:0;left:0;right:0;height:3px;}
.bed-card.good::after{background:#10b981;}
.bed-card.warning::after{background:#f59e0b;}
.bed-card.critical::after{background:#ef4444;}
.bed-ward{font-size:11px;font-weight:700;color:var(--text);margin-bottom:10px;}
.bed-numbers{display:flex;align-items:baseline;gap:4px;margin-bottom:7px;}
.bed-avail{font-family:'JetBrains Mono',monospace;font-size:26px;font-weight:700;}
.bed-avail.good{color:#10b981;}
.bed-avail.warning{color:#f59e0b;}
.bed-avail.critical{color:#ef4444;}
.bed-total{font-size:12px;color:var(--muted);}
.bed-bar{width:100%;height:5px;background:#f1f5f9;border-radius:3px;overflow:hidden;}
.bed-fill{height:100%;border-radius:3px;}
.bed-fill.good{background:#10b981;}
.bed-fill.warning{background:#f59e0b;}
.bed-fill.critical{background:#ef4444;}
.bed-label{font-size:9px;color:var(--muted2);margin-top:5px;}
.ranking-row{display:flex;align-items:center;gap:12px;padding:12px 0;border-bottom:1px solid #f8fafc;}
.ranking-row:last-child{border-bottom:none;}
.rank-num{font-family:'JetBrains Mono',monospace;font-size:18px;font-weight:700;width:34px;text-align:center;color:var(--muted2);}
.rank-name{flex:1;}
.rank-name-main{font-size:13px;font-weight:700;color:var(--text);}
.rank-name-sub{font-size:10px;color:var(--muted2);margin-top:2px;}
.hosp-stars{color:#f59e0b;font-size:12px;}
.hosp-tag{font-size:10px;padding:2px 8px;border-radius:20px;font-weight:700;}
.tag-fast{background:#ecfdf5;color:#059669;border:1px solid #a7f3d0;}
.tag-slow{background:#fff1f2;color:#ef4444;border:1px solid #fecdd3;}
.table-card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:18px;margin-bottom:18px;}
table{width:100%;border-collapse:collapse;}
thead th{font-size:10px;font-weight:700;letter-spacing:0.5px;color:var(--muted2);text-align:left;padding:9px 12px;border-bottom:1px solid var(--border);background:#f8fafc;text-transform:uppercase;}
thead th:first-child{border-radius:7px 0 0 0;}
thead th:last-child{border-radius:0 7px 0 0;}
tbody td{padding:10px 12px;font-size:12px;border-bottom:1px solid #f8fafc;color:var(--text);}
tbody tr:last-child td{border-bottom:none;}
tbody tr:hover td{background:#fafbfc;}
.badge{display:inline-flex;align-items:center;gap:3px;border-radius:20px;padding:3px 9px;font-size:10px;font-weight:700;}
.badge-success{background:#ecfdf5;color:#059669;border:1px solid #a7f3d0;}
.badge-danger{background:#fff1f2;color:#ef4444;border:1px solid #fecdd3;}
.badge-warn{background:#fffbeb;color:#d97706;border:1px solid #fed7aa;}
.badge-blue{background:#eff6ff;color:#3b82f6;border:1px solid #bfdbfe;}
.no-case{display:flex;flex-direction:column;align-items:center;justify-content:center;padding:48px;text-align:center;}
.no-case-icon{font-size:40px;margin-bottom:12px;opacity:0.2;}
.no-case-text{font-size:13px;font-weight:600;color:var(--muted);}
.dispatch-card{background:#ecfdf5;border:1px solid #a7f3d0;border-radius:11px;padding:18px;text-align:center;margin-bottom:14px;}
.dispatch-title{font-family:'Sora',sans-serif;font-size:16px;font-weight:800;color:#059669;margin-bottom:5px;}
.dispatch-sub{font-size:12px;color:var(--muted);}
</style>
</head>
<body>
<div class="toast" id="toast"><span>🏥</span><span id="toast-msg"></span></div>

<div class="sidebar">
  <div class="sb-logo">
    <div class="sb-logo-row">
      <div class="sb-icon">🏥</div>
      <div>
        <div class="sb-name">AcciSense</div>
        <div class="sb-sub">Hospital Portal</div>
      </div>
    </div>
  </div>
  <div class="sb-section">Emergency</div>
  <div class="nav active" onclick="showSection('dashboard',this)"><span class="nav-ic">📊</span> Overview</div>
  <div class="nav" onclick="showSection('incoming',this)"><span class="nav-ic">🚑</span> Incoming Case <span class="nav-badge" id="live-badge" style="display:none">!</span></div>
  <div class="nav" onclick="showSection('analytics',this)"><span class="nav-ic">📈</span> Analytics</div>
  <div class="nav" onclick="showSection('ranking',this)"><span class="nav-ic">🏆</span> Rankings</div>
  <div class="sb-section">Hospital</div>
  <div class="nav" onclick="showSection('beds',this)"><span class="nav-ic">🛏️</span> Bed Status</div>
  <div class="nav" onclick="showSection('fleet',this)"><span class="nav-ic">🚑</span> Fleet Status</div>
  <div class="nav" onclick="showSection('staff',this)"><span class="nav-ic">👨‍⚕️</span> Staff On Duty</div>
  <div class="nav" onclick="showSection('icu',this)"><span class="nav-ic">🩺</span> ICU Monitor</div>
  <div class="sb-bottom">
    <div class="sb-user">
      <div class="av">🏥</div>
      <div>
        <div class="un">Emergency Unit</div>
        <div class="ur">Unit-001 · Hospital</div>
      </div>
    </div>
  </div>
</div>

<div class="topbar">
  <div class="tb-breadcrumb">
    <span class="tb-home">AcciSense</span>
    <span class="tb-sep">/</span>
    <span class="tb-page" id="tb-page-name">Overview</span>
  </div>
  <div class="tb-right">
    <div class="live-pill"><div class="lp-dot"></div>LIVE</div>
    <div class="tb-time" id="clock">--:--:--</div>
    <a href="/" class="tb-logout">← Logout</a>
  </div>
</div>

<div class="main">

  <!-- DASHBOARD -->
  <div id="section-dashboard">
    <div class="page-hdr"><div class="page-title">Hospital Overview</div><div class="page-sub">Monitor incoming cases and manage emergency response</div></div>
    <div class="stats-row">
      <div class="stat-card green"><div class="stat-top"><div class="stat-label">Accepted Today</div><div class="stat-icon-box" style="background:#ecfdf5">✅</div></div><div class="stat-value" id="st-accepted" style="color:#10b981">0</div><div class="stat-sub">Cases dispatched</div></div>
      <div class="stat-card red"><div class="stat-top"><div class="stat-label">Declined Today</div><div class="stat-icon-box" style="background:#fff1f2">❌</div></div><div class="stat-value" id="st-declined" style="color:#ef4444">0</div><div class="stat-sub">Passed to next</div></div>
      <div class="stat-card amber"><div class="stat-top"><div class="stat-label">Avg Response</div><div class="stat-icon-box" style="background:#fffbeb">⏱️</div></div><div class="stat-value" id="st-avgtime" style="color:#f59e0b">--</div><div class="stat-sub">Seconds avg</div></div>
      <div class="stat-card blue"><div class="stat-top"><div class="stat-label">Active Incidents</div><div class="stat-icon-box" style="background:#eff6ff">🚨</div></div><div class="stat-value" id="st-active" style="color:#3b82f6">0</div><div class="stat-sub">Requiring action</div></div>
    </div>
    <div id="dash-alert-hosp"></div>
    <div class="charts-row">
      <div class="chart-card"><div class="chart-title"><span>📊</span> Accept vs Decline</div><canvas id="acceptDoughnut"></canvas></div>
      <div class="chart-card"><div class="chart-title"><span>📉</span> Hospital Decline Rate (%)</div><canvas id="declineBar"></canvas></div>
    </div>
  </div>

  <!-- INCOMING -->
  <div id="section-incoming" style="display:none">
    <div class="page-hdr"><div class="page-title">Incoming Case</div><div class="page-sub">Review and respond to incoming accident cases</div></div>
    <div id="incoming-detail"></div>
  </div>

  <!-- ANALYTICS -->
  <div id="section-analytics" style="display:none">
    <div class="page-hdr"><div class="page-title">Hospital Analytics</div><div class="page-sub">Performance metrics and response analysis</div></div>
    <div class="stats-row">
      <div class="stat-card blue"><div class="stat-top"><div class="stat-label">Total Requests</div><div class="stat-icon-box" style="background:#eff6ff">📋</div></div><div class="stat-value" id="an-total">0</div><div class="stat-sub">All hospitals</div></div>
      <div class="stat-card green"><div class="stat-top"><div class="stat-label">Accept Rate</div><div class="stat-icon-box" style="background:#ecfdf5">✅</div></div><div class="stat-value" id="an-rate" style="color:#10b981">--%</div><div class="stat-sub">Overall percentage</div></div>
      <div class="stat-card green"><div class="stat-top"><div class="stat-label">Fastest Response</div><div class="stat-icon-box" style="background:#ecfdf5">⚡</div></div><div class="stat-value" id="an-fast" style="color:#10b981">--s</div><div class="stat-sub">Best time today</div></div>
      <div class="stat-card purple"><div class="stat-top"><div class="stat-label">Hospitals Active</div><div class="stat-icon-box" style="background:#faf5ff">🏥</div></div><div class="stat-value" id="an-active">5</div><div class="stat-sub">In network</div></div>
    </div>
    <div class="charts-row">
      <div class="chart-card"><div class="chart-title"><span>📊</span> Accepted vs Declined</div><canvas id="an-acceptbar"></canvas></div>
      <div class="chart-card"><div class="chart-title"><span>⏱️</span> Avg Response Time (s)</div><canvas id="an-rtbar"></canvas></div>
    </div>
    <div class="charts-row">
      <div class="chart-card"><div class="chart-title"><span>🏥</span> Accept Rate (%)</div><canvas id="an-rateDonut"></canvas></div>
      <div class="chart-card"><div class="chart-title"><span>📉</span> Decline % by Hospital</div><canvas id="an-declinebar"></canvas></div>
    </div>
    <div class="table-card">
      <div style="font-size:12px;font-weight:700;color:var(--text);margin-bottom:14px;display:flex;align-items:center;gap:7px;"><span>📋</span> Hospital Performance</div>
      <table>
        <thead><tr><th>Hospital</th><th>Total</th><th>Accepted</th><th>Declined</th><th>Avg Response</th><th>Decline Rate</th><th>Rating</th></tr></thead>
        <tbody id="an-tbody"><tr><td colspan="7" style="text-align:center;color:var(--muted2);padding:28px;font-size:12px;">No data yet</td></tr></tbody>
      </table>
    </div>
  </div>

  <!-- RANKING -->
  <div id="section-ranking" style="display:none">
    <div class="page-hdr"><div class="page-title">Hospital Rankings</div><div class="page-sub">Performance-based leaderboard</div></div>
    <div class="alert-card">
      <div style="font-size:12px;font-weight:700;color:var(--text);margin-bottom:14px;display:flex;align-items:center;gap:7px;"><span>🏆</span> Hospital Performance Rankings</div>
      <div id="ranking-list"></div>
    </div>
  </div>

  <!-- BED STATUS -->
  <div id="section-beds" style="display:none">
    <div class="page-hdr"><div class="page-title">Bed Status</div><div class="page-sub">Real-time bed availability across all wards</div></div>
    <div class="stats-row">
      <div class="stat-card blue"><div class="stat-top"><div class="stat-label">Total Beds</div><div class="stat-icon-box" style="background:#eff6ff">🛏️</div></div><div class="stat-value" id="bed-total">0</div><div class="stat-sub">Hospital capacity</div></div>
      <div class="stat-card red"><div class="stat-top"><div class="stat-label">Occupied</div><div class="stat-icon-box" style="background:#fff1f2">🔴</div></div><div class="stat-value" id="bed-occ" style="color:#ef4444">0</div><div class="stat-sub">Currently in use</div></div>
      <div class="stat-card green"><div class="stat-top"><div class="stat-label">Available</div><div class="stat-icon-box" style="background:#ecfdf5">🟢</div></div><div class="stat-value" id="bed-avail" style="color:#10b981">0</div><div class="stat-sub">Ready for patients</div></div>
      <div class="stat-card amber"><div class="stat-top"><div class="stat-label">Occupancy Rate</div><div class="stat-icon-box" style="background:#fffbeb">📊</div></div><div class="stat-value" id="bed-rate" style="color:#f59e0b">0%</div><div class="stat-sub">Capacity used</div></div>
    </div>
    <div class="bed-grid" id="bed-grid"></div>
    <div class="charts-row">
      <div class="chart-card"><div class="chart-title"><span>🛏️</span> Bed Occupancy by Ward</div><canvas id="bedBar"></canvas></div>
      <div class="chart-card"><div class="chart-title"><span>📊</span> Available vs Occupied</div><canvas id="bedDoughnut"></canvas></div>
    </div>
  </div>

  <!-- FLEET STATUS -->
  <div id="section-fleet" style="display:none">
    <div class="page-hdr"><div class="page-title">Fleet Status</div><div class="page-sub">Emergency vehicle availability and tracking</div></div>
    <div class="stats-row">
      <div class="stat-card blue"><div class="stat-top"><div class="stat-label">Total Fleet</div><div class="stat-icon-box" style="background:#eff6ff">🚑</div></div><div class="stat-value" id="fl-total">0</div><div class="stat-sub">All vehicles</div></div>
      <div class="stat-card green"><div class="stat-top"><div class="stat-label">Available</div><div class="stat-icon-box" style="background:#ecfdf5">✅</div></div><div class="stat-value" id="fl-avail" style="color:#10b981">0</div><div class="stat-sub">Ready to deploy</div></div>
      <div class="stat-card blue"><div class="stat-top"><div class="stat-label">On Mission</div><div class="stat-icon-box" style="background:#eff6ff">🚑</div></div><div class="stat-value" id="fl-disp" style="color:#3b82f6">0</div><div class="stat-sub">Dispatched</div></div>
      <div class="stat-card amber"><div class="stat-top"><div class="stat-label">Maintenance</div><div class="stat-icon-box" style="background:#fffbeb">🔧</div></div><div class="stat-value" id="fl-maint" style="color:#f59e0b">0</div><div class="stat-sub">Under service</div></div>
    </div>
    <div class="table-card">
      <div style="font-size:12px;font-weight:700;color:var(--text);margin-bottom:14px;display:flex;align-items:center;gap:7px;"><span>🚑</span> Ambulance Fleet</div>
      <table>
        <thead><tr><th>Vehicle ID</th><th>Type</th><th>Driver</th><th>Location</th><th>Last Trip</th><th>Status</th></tr></thead>
        <tbody id="fl-tbody"></tbody>
      </table>
    </div>
  </div>

  <!-- STAFF -->
  <div id="section-staff" style="display:none">
    <div class="page-hdr"><div class="page-title">Staff On Duty</div><div class="page-sub">Current medical staff on emergency duty</div></div>
    <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:14px;">
      <div class="alert-card"><div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;"><div style="font-size:28px;">👨‍⚕️</div><div><div style="font-weight:700;font-size:13px;">Dr. Arjun Mehta</div><div style="font-size:10px;color:var(--muted);">Emergency Surgeon</div></div></div><span class="badge badge-success">● On Duty</span></div>
      <div class="alert-card"><div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;"><div style="font-size:28px;">👩‍⚕️</div><div><div style="font-weight:700;font-size:13px;">Dr. Priya Sharma</div><div style="font-size:10px;color:var(--muted);">Trauma Specialist</div></div></div><span class="badge badge-success">● On Duty</span></div>
      <div class="alert-card"><div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;"><div style="font-size:28px;">👨‍⚕️</div><div><div style="font-weight:700;font-size:13px;">Dr. Ravi Kumar</div><div style="font-size:10px;color:var(--muted);">Cardiologist</div></div></div><span class="badge badge-warn">⚠ Break</span></div>
      <div class="alert-card"><div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;"><div style="font-size:28px;">👩‍⚕️</div><div><div style="font-weight:700;font-size:13px;">Dr. Sunita Patel</div><div style="font-size:10px;color:var(--muted);">Neurologist</div></div></div><span class="badge badge-success">● On Duty</span></div>
      <div class="alert-card"><div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;"><div style="font-size:28px;">👨‍⚕️</div><div><div style="font-weight:700;font-size:13px;">Dr. Anil Reddy</div><div style="font-size:10px;color:var(--muted);">Orthopedic Surgeon</div></div></div><span class="badge badge-success">● On Duty</span></div>
      <div class="alert-card"><div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;"><div style="font-size:28px;">👩‍⚕️</div><div><div style="font-weight:700;font-size:13px;">Dr. Kavya Nair</div><div style="font-size:10px;color:var(--muted);">ICU Specialist</div></div></div><span class="badge badge-danger">● Off Duty</span></div>
    </div>
  </div>

  <!-- ICU MONITOR -->
  <div id="section-icu" style="display:none">
    <div class="page-hdr"><div class="page-title">ICU Monitor</div><div class="page-sub">Intensive care unit patient monitoring</div></div>
    <div class="stats-row">
      <div class="stat-card purple"><div class="stat-top"><div class="stat-label">ICU Total</div><div class="stat-icon-box" style="background:#faf5ff">🩺</div></div><div class="stat-value">20</div><div class="stat-sub">Emergency ICU capacity</div></div>
      <div class="stat-card red"><div class="stat-top"><div class="stat-label">ICU Occupied</div><div class="stat-icon-box" style="background:#fff1f2">🔴</div></div><div class="stat-value" style="color:#ef4444">14</div><div class="stat-sub">Critical patients</div></div>
      <div class="stat-card green"><div class="stat-top"><div class="stat-label">ICU Available</div><div class="stat-icon-box" style="background:#ecfdf5">🟢</div></div><div class="stat-value" style="color:#10b981">6</div><div class="stat-sub">Ready for emergency</div></div>
      <div class="stat-card red"><div class="stat-top"><div class="stat-label">Critical Patients</div><div class="stat-icon-box" style="background:#fff1f2">⚠️</div></div><div class="stat-value" style="color:#ef4444">3</div><div class="stat-sub">Immediate attention</div></div>
    </div>
    <div class="table-card">
      <div style="font-size:12px;font-weight:700;margin-bottom:14px;display:flex;align-items:center;gap:7px;"><span>🩺</span> ICU Patient Status</div>
      <table>
        <thead><tr><th>Bed #</th><th>Patient ID</th><th>Condition</th><th>Admitted</th><th>Doctor</th><th>Status</th></tr></thead>
        <tbody>
          <tr><td><span style="font-family:'JetBrains Mono',monospace;font-size:11px;">ICU-01</span></td><td>P-2401</td><td>Head Trauma</td><td>09:15 AM</td><td>Dr. Arjun Mehta</td><td><span class="badge badge-danger">Critical</span></td></tr>
          <tr><td><span style="font-family:'JetBrains Mono',monospace;font-size:11px;">ICU-02</span></td><td>P-2402</td><td>Cardiac Arrest</td><td>10:30 AM</td><td>Dr. Ravi Kumar</td><td><span class="badge badge-danger">Critical</span></td></tr>
          <tr><td><span style="font-family:'JetBrains Mono',monospace;font-size:11px;">ICU-03</span></td><td>P-2403</td><td>Spinal Injury</td><td>08:45 AM</td><td>Dr. Anil Reddy</td><td><span class="badge badge-warn">Serious</span></td></tr>
          <tr><td><span style="font-family:'JetBrains Mono',monospace;font-size:11px;">ICU-04</span></td><td>P-2404</td><td>Internal Bleeding</td><td>11:00 AM</td><td>Dr. Priya Sharma</td><td><span class="badge badge-warn">Serious</span></td></tr>
          <tr><td><span style="font-family:'JetBrains Mono',monospace;font-size:11px;">ICU-05</span></td><td>P-2405</td><td>Fractures</td><td>07:20 AM</td><td>Dr. Anil Reddy</td><td><span class="badge badge-blue">Stable</span></td></tr>
        </tbody>
      </table>
    </div>
  </div>

</div>

<script>
let CS={},lastCases=0;
let acceptDChart,declineBarC,anAccBar,anRtBar,anRateD,anDecBar,bedBarC,bedDC;
let bedChartInit=false,anChartInit=false;

setInterval(()=>document.getElementById('clock').textContent=new Date().toLocaleTimeString('en-IN',{hour12:false}),1000);

function showSection(name,el){
  document.querySelectorAll('[id^="section-"]').forEach(s=>s.style.display='none');
  document.getElementById('section-'+name).style.display='block';
  document.querySelectorAll('.nav').forEach(n=>n.classList.remove('active'));
  if(el)el.classList.add('active');
  const names={dashboard:'Overview',incoming:'Incoming Case',analytics:'Analytics',ranking:'Rankings',beds:'Bed Status',fleet:'Fleet Status',staff:'Staff On Duty',icu:'ICU Monitor'};
  document.getElementById('tb-page-name').textContent=names[name]||name;
  if(name==='analytics')renderAnalytics();
  if(name==='ranking')renderRanking();
  if(name==='beds')renderBeds();
  if(name==='fleet')renderFleet();
}

function toast(msg){const t=document.getElementById('toast');document.getElementById('toast-msg').textContent=msg;t.classList.add('show');setTimeout(()=>t.classList.remove('show'),6000);}
function fmtTime(iso){return iso?new Date(iso).toLocaleTimeString('en-IN',{hour12:false}):'--';}
function getVI(p){const VT={TN01:{type:"Sedan",icon:"🚗"},TN02:{type:"SUV",icon:"🚙"},TN03:{type:"Truck",icon:"🚛"},MH:{type:"Motorcycle",icon:"🏍️"},KA:{type:"Auto-Rickshaw",icon:"🛺"},AP:{type:"Bus",icon:"🚌"},TS:{type:"Van",icon:"🚐"},DL:{type:"Hatchback",icon:"🚘"},GJ:{type:"Pickup",icon:"🛻"},WB:{type:"Taxi",icon:"🚕"}};for(let k in VT){if(p.toUpperCase().startsWith(k))return VT[k];}return{type:"Vehicle",icon:"🚗"};}

function buildHospCard(s){
  if(!s.current_case)return'<div class="no-case"><div class="no-case-icon">🏥</div><div class="no-case-text">No incoming case — System ready</div></div>';
  const c=s.current_case;
  const mapUrl=`https://maps.google.com/maps?q=${c.latitude},${c.longitude}&z=16&output=embed`;
  const idx=s.hospital_index;
  const HOSPS=[{name:"Apollo Emergency Center",dist:"1.2 km"},{name:"City Care Hospital",dist:"2.4 km"},{name:"Metro Trauma Hospital",dist:"3.1 km"},{name:"Green Cross Medical",dist:"4.8 km"},{name:"National Emergency Hospital",dist:"6.2 km"}];
  const thisH=idx<HOSPS.length?HOSPS[idx]:null;
  const showBtns=s.case_status==='pending'&&thisH;
  const v1=getVI(c.vehicle_1),v2=getVI(c.vehicle_2);
  const stCls=s.case_status==='pending'?'pending':s.case_status==='ambulance_dispatched'?'dispatched':'no_hospital';
  let dispHtml='';
  if(s.case_status==='ambulance_dispatched'){dispHtml=`<div class="dispatch-card"><div class="dispatch-title">🚑 Ambulance Dispatched</div><div class="dispatch-sub">Alert: ${fmtTime(s.alert_time)} &nbsp;·&nbsp; Accepted: ${fmtTime(s.accept_time)}</div></div>`;}
  return`<div class="alert-card ${s.case_status==='pending'?'emergency':''}">
    <div class="alert-header">
      ${s.case_status==='pending'?'<div class="emergency-pill">🚨 INCOMING CASE</div>':'<div class="emergency-pill" style="background:#ecfdf5;border-color:#a7f3d0;color:#059669;animation:none">✅ RESOLVED</div>'}
      <div class="alert-title">Emergency Accident Response Required</div>
      <div class="alert-status status-${stCls}">${s.case_status.replace(/_/g,' ').toUpperCase()}</div>
    </div>
    ${dispHtml}
    <div class="info-grid">
      <div class="info-field"><div class="ifl">Hospital Contacted</div><div class="ifv">${thisH?thisH.name:HOSPS[idx-1]?.name||'N/A'}</div></div>
      <div class="info-field"><div class="ifl">Distance</div><div class="ifv">${thisH?thisH.dist:'N/A'}</div></div>
      <div class="info-field"><div class="ifl">Alert Time</div><div class="ifv">${fmtTime(s.alert_time)}</div></div>
      <div class="info-field"><div class="ifl">Vehicles</div><div class="ifv">${v1.icon} ${c.vehicle_1} / ${v2.icon} ${c.vehicle_2}</div></div>
      <div class="info-field"><div class="ifl">GPS Location</div><div class="ifv" style="font-family:'JetBrains Mono',monospace;font-size:11px;">${c.latitude}, ${c.longitude}</div></div>
      <div class="info-field"><div class="ifl">Police Assigned</div><div class="ifv" style="color:#f59e0b;font-size:12px;">🚓 ${s.selected_police}</div></div>
    </div>
    <iframe class="map-iframe" src="${mapUrl}" allowfullscreen loading="lazy" style="margin-bottom:${showBtns?'14px':'0'}"></iframe>
    ${showBtns?`<div class="action-btns">
      <form action="/accept" method="post" style="flex:1"><button class="btn-accept" type="submit">✅ Accept & Dispatch Ambulance</button></form>
      <form action="/decline" method="post" style="flex:1"><button class="btn-decline" type="submit">❌ Decline — Next Hospital</button></form>
    </div>`:''}
  </div>`;
}

function initDashCharts(){
  Chart.defaults.font.family='Sora';Chart.defaults.color='#64748b';Chart.defaults.font.size=11;
  acceptDChart=new Chart(document.getElementById('acceptDoughnut'),{type:'doughnut',data:{labels:['Accepted','Declined'],datasets:[{data:[0,0],backgroundColor:['rgba(16,185,129,0.7)','rgba(239,68,68,0.7)'],borderWidth:0}]},options:{plugins:{legend:{position:'bottom',labels:{boxWidth:8,padding:8}}},cutout:'62%',animation:false}});
  declineBarC=new Chart(document.getElementById('declineBar'),{type:'bar',data:{labels:['Apollo','City Care','Metro','Green Cross','National'],datasets:[{label:'Decline %',data:[0,0,0,0,0],backgroundColor:'rgba(239,68,68,0.4)',borderColor:'#ef4444',borderWidth:1,borderRadius:5}]},options:{plugins:{legend:{display:false}},indexAxis:'y',scales:{x:{beginAtZero:true,max:100,grid:{color:'rgba(0,0,0,0.04)'}},y:{grid:{display:false}}},animation:false}});
}

async function renderAnalytics(){
  const r=await fetch('/api/hospital-stats');const stats=await r.json();
  const names=stats.map(s=>s.name.split(' ')[0]);
  const acc=stats.map(s=>s.accepted),dec=stats.map(s=>s.declined);
  const avgT=stats.map(s=>s.avg_time),decP=stats.map(s=>s.decline_pct);
  const totR=stats.reduce((a,s)=>a+s.total,0);
  const totA=stats.reduce((a,s)=>a+s.accepted,0);
  const rate=totR>0?Math.round(totA/totR*100):0;
  const times=stats.flatMap(s=>s.avg_time>0?[s.avg_time]:[]);
  const fast=times.length?Math.min(...times):null;
  document.getElementById('an-total').textContent=totR;
  document.getElementById('an-rate').textContent=rate+'%';
  document.getElementById('an-fast').textContent=fast?fast+'s':'--s';
  if(!anChartInit){
    anChartInit=true;
    anAccBar=new Chart(document.getElementById('an-acceptbar'),{type:'bar',data:{labels:names,datasets:[{label:'Accepted',data:acc,backgroundColor:'rgba(16,185,129,0.5)',borderColor:'#10b981',borderWidth:1,borderRadius:5},{label:'Declined',data:dec,backgroundColor:'rgba(239,68,68,0.5)',borderColor:'#ef4444',borderWidth:1,borderRadius:5}]},options:{plugins:{legend:{labels:{boxWidth:8,font:{size:10}}}},scales:{y:{beginAtZero:true,grid:{color:'rgba(0,0,0,0.04)'}},x:{grid:{display:false}}},animation:false}});
    anRtBar=new Chart(document.getElementById('an-rtbar'),{type:'bar',data:{labels:names,datasets:[{label:'Avg Response(s)',data:avgT,backgroundColor:'rgba(59,130,246,0.5)',borderColor:'#3b82f6',borderWidth:1,borderRadius:5}]},options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,grid:{color:'rgba(0,0,0,0.04)'}},x:{grid:{display:false}}},animation:false}});
    const accRates=stats.map(s=>s.total>0?Math.round(s.accepted/s.total*100):0);
    anRateD=new Chart(document.getElementById('an-rateDonut'),{type:'doughnut',data:{labels:names,datasets:[{data:accRates,backgroundColor:['#10b981','#3b82f6','#f59e0b','#ef4444','#8b5cf6'],borderWidth:0}]},options:{plugins:{legend:{position:'bottom',labels:{boxWidth:8,font:{size:10}}}},cutout:'55%',animation:false}});
    anDecBar=new Chart(document.getElementById('an-declinebar'),{type:'bar',data:{labels:names,datasets:[{label:'Decline %',data:decP,backgroundColor:'rgba(239,68,68,0.4)',borderColor:'#ef4444',borderWidth:1,borderRadius:5}]},options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,max:100,grid:{color:'rgba(0,0,0,0.04)'}},x:{grid:{display:false}}},animation:false}});
  }else{anAccBar.data.datasets[0].data=acc;anAccBar.data.datasets[1].data=dec;anAccBar.update();anRtBar.data.datasets[0].data=avgT;anRtBar.update();anDecBar.data.datasets[0].data=decP;anDecBar.update();}
  const tbody=document.getElementById('an-tbody');
  tbody.innerHTML=stats.map((s,i)=>{const bc=s.decline_pct>50?'badge-danger':s.decline_pct>20?'badge-warn':'badge-success';const stars='★'.repeat(Math.max(1,5-i));return`<tr><td style="font-weight:700">${s.name}</td><td>${s.total}</td><td style="color:#10b981;font-weight:700">${s.accepted}</td><td style="color:#ef4444;font-weight:700">${s.declined}</td><td><span style="font-family:'JetBrains Mono',monospace;font-size:11px;">${s.avg_time>0?s.avg_time+'s':'--'}</span></td><td><span class="badge ${bc}">${s.decline_pct}%</span></td><td style="color:#f59e0b">${stars}</td></tr>`;}).join('');
}

async function renderRanking(){
  const r=await fetch('/api/hospital-stats');const stats=await r.json();
  const sorted=[...stats].sort((a,b)=>{if(a.avg_time===0&&b.avg_time===0)return b.accepted-a.accepted;if(a.avg_time===0)return 1;if(b.avg_time===0)return -1;return a.avg_time-b.avg_time;});
  const maxD=Math.max(...stats.map(s=>s.decline_pct),0);
  const medals=['🥇','🥈','🥉'];
  document.getElementById('ranking-list').innerHTML=sorted.map((h,i)=>{
    const stars='★'.repeat(Math.max(1,5-i));
    const isFastest=i===0&&h.avg_time>0;
    const isMostD=h.decline_pct===maxD&&maxD>0;
    return`<div class="ranking-row"><div class="rank-num">${medals[i]||i+1}</div><div class="rank-name"><div class="rank-name-main">${h.name}</div><div class="rank-name-sub">Accepted: ${h.accepted} · Declined: ${h.declined} · Avg: ${h.avg_time>0?h.avg_time+'s':'--'}</div></div><div style="display:flex;gap:5px;align-items:center">${isFastest?'<span class="hosp-tag tag-fast">⚡ Fastest</span>':''}${isMostD?'<span class="hosp-tag tag-slow">⚠ High Decline</span>':''}</div><div class="hosp-stars">${stars}</div><div style="font-family:'JetBrains Mono',monospace;font-size:12px;color:${h.decline_pct>50?'#ef4444':'#10b981'};min-width:44px;text-align:right;font-weight:700">${h.decline_pct}%</div></div>`;
  }).join('');
}

async function renderBeds(){
  const r=await fetch('/api/bed-status');const beds=await r.json();
  let tot=0,occ=0,avail=0;
  beds.forEach(b=>{tot+=b.total;occ+=b.occupied;avail+=b.available;});
  document.getElementById('bed-total').textContent=tot;
  document.getElementById('bed-occ').textContent=occ;
  document.getElementById('bed-avail').textContent=avail;
  document.getElementById('bed-rate').textContent=Math.round(occ/tot*100)+'%';
  document.getElementById('bed-grid').innerHTML=beds.map(b=>{const pct=Math.round(b.occupied/b.total*100);return`<div class="bed-card ${b.status}"><div class="bed-ward">${b.ward}</div><div class="bed-numbers"><div class="bed-avail ${b.status}">${b.available}</div><div class="bed-total">/ ${b.total}</div></div><div class="bed-bar"><div class="bed-fill ${b.status}" style="width:${pct}%"></div></div><div class="bed-label">${pct}% occupied · ${b.available} free</div></div>`;}).join('');
  if(!bedChartInit){
    bedChartInit=true;
    bedBarC=new Chart(document.getElementById('bedBar'),{type:'bar',data:{labels:beds.map(b=>b.ward.replace(' Ward','').replace(' Unit','')),datasets:[{label:'Occupied',data:beds.map(b=>b.occupied),backgroundColor:'rgba(239,68,68,0.5)',borderColor:'#ef4444',borderWidth:1,borderRadius:4},{label:'Available',data:beds.map(b=>b.available),backgroundColor:'rgba(16,185,129,0.5)',borderColor:'#10b981',borderWidth:1,borderRadius:4}]},options:{plugins:{legend:{labels:{boxWidth:8,font:{size:10}}}},scales:{y:{beginAtZero:true,grid:{color:'rgba(0,0,0,0.04)'}},x:{grid:{display:false},ticks:{maxRotation:45,font:{size:9}}}},animation:false}});
    bedDC=new Chart(document.getElementById('bedDoughnut'),{type:'doughnut',data:{labels:['Occupied','Available'],datasets:[{data:[occ,avail],backgroundColor:['rgba(239,68,68,0.6)','rgba(16,185,129,0.6)'],borderWidth:0}]},options:{plugins:{legend:{position:'bottom',labels:{boxWidth:8,font:{size:11}}}},cutout:'62%',animation:false}});
  }
}

async function renderFleet(){
  const r=await fetch('/api/fleet-status');const data=await r.json();
  let av=0,di=0,ma=0;
  data.forEach(f=>{if(f.status==='available')av++;else if(f.status==='dispatched')di++;else ma++;});
  document.getElementById('fl-total').textContent=data.length;
  document.getElementById('fl-avail').textContent=av;
  document.getElementById('fl-disp').textContent=di;
  document.getElementById('fl-maint').textContent=ma;
  document.getElementById('fl-tbody').innerHTML=data.map(f=>{const bc=f.status==='available'?'badge-success':f.status==='dispatched'?'badge-blue':'badge-warn';const lbl=f.status==='available'?'✅ Available':f.status==='dispatched'?'🚑 Dispatched':'🔧 Maintenance';return`<tr><td><strong style="font-family:'JetBrains Mono',monospace;font-size:11px;">${f.id}</strong></td><td>${f.type}</td><td>${f.driver}</td><td>${f.location}</td><td>${f.last_trip}</td><td><span class="badge ${bc}">${lbl}</span></td></tr>`;}).join('');
}

async function poll(){
  try{
    const r=await fetch('/api/state');const s=await r.json();CS=s;
    const hs=s.hospital_stats||{};
    let totA=0,totD=0,totT=0,totN=0;
    Object.values(hs).forEach(h=>{totA+=h.accepted;totD+=h.declined;h.times.forEach(t=>{totT+=t;totN++;});});
    document.getElementById('st-accepted').textContent=totA;
    document.getElementById('st-declined').textContent=totD;
    document.getElementById('st-avgtime').textContent=totN>0?Math.round(totT/totN):'--';
    document.getElementById('st-active').textContent=s.case_status==='pending'?1:0;
    const b=document.getElementById('live-badge');
    b.style.display=s.case_status==='pending'?'inline':'none';
    if(s.current_case&&s.total_cases>lastCases&&s.case_status==='pending'){toast('Emergency case incoming — immediate response required');lastCases=s.total_cases;}
    if(s.total_cases>0)lastCases=s.total_cases;
    document.getElementById('dash-alert-hosp').innerHTML=buildHospCard(s);
    document.getElementById('incoming-detail').innerHTML=buildHospCard(s);
    if(acceptDChart){acceptDChart.data.datasets[0].data=[totA,totD];acceptDChart.update();}
    if(declineBarC){const statsR=await fetch('/api/hospital-stats');const statsD=await statsR.json();declineBarC.data.datasets[0].data=statsD.map(st=>st.decline_pct);declineBarC.update();}
  }catch(e){}
}

initDashCharts();poll();setInterval(poll,2000);
</script>
</body>
</html>"""