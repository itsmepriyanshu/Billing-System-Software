# Billing-System-Software
# Billing System Software

A modern **Python Tkinter Desktop Billing System** designed for retail and stationery stores. It includes inventory management, customer records, automatic discount and VAT calculations, professional PDF receipt generation, and sales history—all in an easy-to-use desktop application.

## Features

* 🧾 Professional receipt generation (PDF)
* 📦 Product and inventory management
* 👤 Customer information management
* 💰 Automatic discount calculation
* 🏛️ 13% VAT calculation
* 📊 Sales history and records
* 🖨️ Thermal printer-ready receipts
* 🖥️ Clean Tkinter desktop interface
* 💾 Local database support (SQLite)

## Tech Stack

* **Python 3**
* **Tkinter** (Desktop GUI)
* **SQLite** (Local Database)
* **ReportLab** (PDF Receipts)
* **PyInstaller** (Windows Executable)

## Project Structure

```text
Billing-System-Software/
│── app.py
│── database.db
│── receipts/
│── assets/
│── requirements.txt
└── README.md
```

## Installation

Clone the repository:

```bash
git clone https://github.com/itsmepriyanshu/Billing-System-Software.git
cd Billing-System-Software
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Run the application:

```bash
python app.py
```

## Build Windows Executable

Create a standalone `.exe`:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed app.py
```

The executable will be available inside the `dist` folder.

## Receipt Features

Each receipt includes:

* Store name
* Sale ID
* Date & Time
* Customer details
* Itemized products
* Quantity, Rate & Amount
* Discount
* VAT (13%)
* Grand Total
* Payment method
* Thank-you message

## Deployment

This project is currently a **Tkinter desktop application** 

## Future Improvements

* Barcode scanner support
* QR code on receipts
* Multiple payment methods
* User authentication
* Cloud database synchronization
* Sales analytics dashboard
* Multi-store support

## License

This project is licensed under the Serab License.
