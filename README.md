# AcciSense: A Deep Visual Learning Framework for Instant Road Accident Detection and Alert Generation.

## Overview

AcciSense is an AI-powered real-time road accident detection and emergency alert generation system that utilizes existing CCTV surveillance cameras to monitor roads, highways, residential areas, and public spaces.

The system automatically detects road accidents using deep learning models, extracts vehicle number plate information, and generates instant emergency alerts to nearby hospitals and police stations. By eliminating the dependency on manual reporting, AcciSense significantly reduces emergency response time and improves public safety.


## Problem Statement

Road accidents often result in severe injuries and fatalities because emergency services are not informed immediately. Existing systems rely heavily on eyewitness reporting, leading to delays in ambulance dispatch and police response.

AcciSense addresses this challenge by providing an automated accident detection and notification framework capable of operating continuously using existing CCTV infrastructure.


## Objectives

- Detect road accidents automatically using CCTV footage.
- Reduce emergency response time.
- Send instant alerts to nearby hospitals and police stations.
- Improve ambulance dispatch efficiency.
- Minimize false alerts through multi-frame analysis.


## Key Features

- Real-time accident detection using deep learning.
- Vehicle number plate recognition.
- Automatic emergency alert generation.
- Hospital notification system.
- Police control room notification.
- Automated alert forwarding to alternate hospitals when ambulance services are unavailable.
- Voice-based emergency alert generation.
- Continuous monitoring using CCTV cameras.


## System Workflow

1. CCTV video feed is continuously monitored.
2. Deep learning models detect road accidents.
3. Vehicle number plate information is extracted.
4. Emergency alerts are generated automatically.
5. Nearby hospitals receive ambulance requests.
6. If a hospital declines due to ambulance unavailability, the alert is forwarded to the next nearest hospital.
7. Police control room receives accident details.
8. Emergency services are dispatched to the accident location.


## Technologies Used

- Python
- OpenCV
- YOLO-based Object Detection
- OCR for Number Plate Recognition
- FastAPI
- Uvicorn
- Computer Vision
- Deep Learning


## Project Structure

ACCISENSE/
│
├── src/
├── config/
│   └── places.json
├── screenshots/
├── .gitignore
├── requirements.txt
└── README.md


## How to Run

Step 1: Start Accident Detection

Open another terminal and run:

py accisense.py --source "clip1.mp4" --accident_model accident.pt --plate_model plate.pt

Step 2: Start the Alert Server

Open a terminal and run:

uvicorn alert_server:app --reload


Step 3: Access the Web Application

After both services start successfully, a localhost URL will be displayed in the terminal.

Open the generated localhost link in your browser to access the AcciSense web interface, monitor accident alerts, and view system updates in real time.

Note : Model files and sample videos are not included in this respository due to size limitations.


## Screenshots

### Accident Detection
![Accident Detection](screenshots/accident_detection.jpeg)

### Terminal Output
![Emergency Alert Generation](screenshots/alert_terminal.jpeg)

### Website Login
![Website Login](screenshots/login.jpeg)

### Hospital Dashboard
![Hospital Dashboard](screenshots/hospital_dashoard.jpeg)

### Police Dashboard
![Police Dashboard](screenshots/police_dashboard.jpeg)


## Applications

- Smart City Infrastructure
- Highway Monitoring Systems
- Urban Traffic Management
- Emergency Response Systems
- Public Safety Monitoring


## Future Enhancements

- Improve accident detection accuracy using advanced AI models and larger datasets.
- Optimize frame processing speed for enhanced real-time performance.
- Develop a dedicated monitoring and reporting application.
- Deploy on edge devices such as Raspberry Pi and NVIDIA Jetson.
- Integrate with smart traffic management systems.


## Conclusion

AcciSense provides an intelligent and automated accident detection and emergency response framework that minimizes reporting delays and improves coordination between hospitals, police departments, and vehicle owners. By leveraging deep learning and computer vision technologies, the system contributes to safer roads and faster emergency assistance.


## Team Members 

- R. Divya
- S. Pavithra
- S. Sreelaya


## Academic Project

Final Year B.Tech. Artificial Intelligence and Data Science Project