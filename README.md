# 📄 Smart Invoice Automation & Reporting System

A lightweight, modular, and developer-friendly web system built with **Python & Flask** to automate invoice generation, report calculations, and Excel-based financial sheet synchronization. Designed using **Clean Architecture** principles to enforce a strict separation of concerns between web routes, business logic, and data persistence.

---

## 🎯 Business Context & Customization

This system was custom-built to solve a specific client's workflow requirements:
* **Tailored Workflow**: Designed to match the exact invoicing lifecycle and calculations required by the business.
* **Custom Reporting**: Generates invoices structured specifically around the client's pre-defined Excel data formats.
* **Generic Template Notice**: All proprietary logos, client names, and confidential business data have been anonymized and replaced with placeholder assets for demonstration purposes.

---

## 🌟 Key Features

* **Invoice Generation**: Automates client invoice creation and exports formatted output (PDF/Word).
* **Excel Data Engine**: Interacts directly with Excel spreadsheets (`.xlsx`) for dynamic data reading, updating, and record persistence without overhead.
* **Modular Architecture**: Built with Flask Blueprints and dedicated layers for business logic and data access.

---

## 🏗️ Architecture & Project Structure

The project follows a **Clean Architecture (Separation of Concerns)** pattern to ensure high maintainability and ease of testing.

### Architectural Breakdown:
* **Presentation Layer (`routes/`)**: Handles HTTP requests, extracts parameters, and delegates work to services.
* **Business Logic Layer (`services/`)**: Contains pure business rules (e.g., total calculations, date formatting, document rendering).
* **Data Access Layer (`data/`)**: Isolates all Direct File/Excel operations using Python libraries (`openpyxl` / `pandas`), keeping storage mechanics completely decoupled from web logic.

---

## 🛠️ Tech Stack

* **Backend**: Python 3.12.10, Flask (Micro-framework, Blueprints)
* **Data & Automation**: Excel Integration (`openpyxl`), Document Processing (`ReportLab` / `python-docx`)
* **Frontend**: HTML5, CSS3, JavaScript (Fetch API / DOM Manipulation)
* **Version Control**: Git & GitHub

---

## 🚀 Quick Start Guide

### Prerequisites
* Python 3.12.10 or higher installed on your machine.

### Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/invoice-automation-system.git
   cd invoice-automation-system
   ```

2. Create and activate a virtual environment:
   ```bash
   # Windows
   python -m venv venv
   venv\Scripts\activate

   # macOS/Linux
   python3 -m venv venv
   source venv/bin/activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Run the application:
   ```bash
   python app.py
   ```

5. Access the application:
   * Open your browser and navigate to `http://127.0.0.1:5000`

---

## 🛡️ Data Privacy & Dummy Data Notice

This repository contains only dummy data and placeholder Excel spreadsheets for demonstration and testing purposes. No sensitive or real business records are stored within this source code.

---

## 📝 License

Distributed under the MIT License. See `LICENSE` for more information.