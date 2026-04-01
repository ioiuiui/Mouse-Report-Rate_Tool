# Mouse Report Rate Tool

A Python-based tool for measuring and analyzing mouse report rate performance using Raw Input API, designed for real-time monitoring and testing scenarios.

---

## 🔍 Overview

This project provides a practical solution for capturing and analyzing mouse report rate behavior in real time.
It is built for performance validation, testing environments, and engineering-level diagnostics.

The tool focuses on **precision, responsiveness, and real-time data visibility**, making it suitable for use cases such as hardware validation, gaming performance testing, and system-level input analysis.

---

## ⚙️ Key Features

* 📡 **Raw Input API Integration**
  Directly captures low-level mouse input data for high accuracy

* 📊 **Real-Time Report Rate Monitoring**
  Measures and visualizes polling rate dynamically

* ⚡ **High-Frequency Event Tracking**
  Supports high polling rate devices (e.g., 1000Hz+)

* 🧪 **Testing-Oriented Design**
  Built for validation scenarios and performance analysis

* 🔧 **Multi-module Utility Support**
  Includes capture tools, device configuration, and supporting scripts

---

## 🧱 Project Structure

```
.
├── Attenuator_control.py        # Attenuator control logic
├── n9010a_capture.py            # Capture integration module
├── report_rate_package_capture_20260209.py  # Main report rate logic
├── device_db.json               # Device configuration
├── VNX_atten64.dll              # External dependency
├── .gitignore
└── README.md
```

---

## 🧠 Technical Highlights

* Uses **Windows Raw Input API** for accurate event capture
* Handles **high-frequency input streams** efficiently
* Designed with **testability and modularity** in mind
* Suitable for **engineering diagnostics and validation workflows**

---

## 🖥️ Environment

* Python
* Windows OS
* PyCharm (recommended)

---

## 🚀 Usage

1. Create a virtual environment:

   ```
   python -m venv .venv
   ```

2. Activate environment:

   ```
   .venv\Scripts\activate
   ```

3. Run the main script:

   ```
   python report_rate_package_capture_20260209.py
   ```

---

## 📌 Notes

* Virtual environment and build artifacts are excluded via `.gitignore`
* This project focuses on functionality and performance testing rather than UI design

---

## 👨‍💻 Author

Developed as part of performance analysis and tool development practice.

---
