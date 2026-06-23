import random
import time

POLICE_STATIONS = [
    "Central Police Station",
    "North Traffic Police Station",
    "City Control Room"
]

def notify_police_station():
    """
    Notify only ONE police station.
    No accept/decline logic.
    """

    station = random.choice(POLICE_STATIONS)

    print(f"\n🚓 Notifying {station}...")
    time.sleep(1)
    print(f"✅ {station} has been informed")

    return station