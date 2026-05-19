# 🏦 Desktop Banking System

A professional, Python-based graphical user interface (GUI) for a local banking system. This desktop application is designed to simulate core banking operations with a focus on clean UI/UX, robust data validation, and dual-state local storage using Microsoft Excel and text-based logging.

## ✨ Features

* **Account Management:** Create new bank accounts with comprehensive KYC details (Name, DOB, Phone, Gender, Account Type, Nominee).
* **Financial Transactions:** Securely process deposits and withdrawals with real-time balance validation.
* **Account Deletion:** Safely close and delete accounts with an active confirmation safeguard.
* **Dual-Storage Backend:** 
  * Maintains the active state database in an Excel file (`.xlsx`).
  * Records an immutable, timestamped transaction history in a log file (`.txt`).
* **Dynamic Transaction History:** View all system transactions formatted cleanly in a scrollable UI data table.
* **Professional UI/UX:** Features a minimalist layout, responsive ghost text (placeholders), hover cursors, and a dynamic status bar for color-coded error/success handling.

## 🛠️ Technology Stack

* **Language:** Python 3
* **GUI Framework:** `tkinter` (Native Python) & `ttk` (Themed Tkinter)
* **Database/Storage:** `openpyxl` (Excel), Standard File I/O (`.txt`)
* **Core Libraries:** `os`, `datetime`

## 📦 Prerequisites

Before running this application, ensure you have Python installed on your system. You will also need to install the `openpyxl` library to handle the Excel database operations.

```bash
pip install openpyxl
