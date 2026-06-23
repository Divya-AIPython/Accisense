import random
import time

HOSPITALS = [
    "City Care Hospital",
    "Apollo Emergency Center",
    "Metro Trauma Hospital"
]

def process_hospital_chain():
    """
    Contact hospitals one by one.
    Stop when one accepts.
    """

    for hospital in HOSPITALS:
        print(f"\n📡 Contacting {hospital}...")
        time.sleep(1)

        accepted = random.random() < 0.6  # 60% chance to accept

        if accepted:
            print(f"✅ {hospital} ACCEPTED the case")
            print(f"🚑 Ambulance dispatched from {hospital}")
            return {
                "status": "accepted",
                "hospital": hospital,
                "ambulance": "dispatched"
            }

        else:
            print(f"❌ {hospital} DECLINED the case")

    print("\n❌ All hospitals declined the case")
    return {
        "status": "failed",
        "hospital": None,
        "ambulance": "not_available"
    }