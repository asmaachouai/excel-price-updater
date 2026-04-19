📊 Excel Price Updater & Visualization Tool

A Python-based automation tool that processes Excel spreadsheets, applies price corrections, and generates visual insights through charts.

🚀 Overview:

This project demonstrates how Python can be used to automate repetitive data processing tasks. It reads an Excel file, applies a pricing rule (discount), writes updated values into a new column, and generates a bar chart for visualization.

The goal is to showcase practical Python skills in:

- Data processing.
- File handling.
- Automation.
- Basic data visualization.

⚙️ Features:

📥 Load Excel .xlsx files.
💰 Apply automated price correction (e.g., 10% discount).
🧮 Validate and process numeric data.
📊 Generate bar chart visualization.
💾 Save processed data to a new file.
🧱 Modular and maintainable code structure.

🧱 Project Structure:

excel-price-updater/
│
├── main.py                  # Main application logic
├── transactions.xlsx        # Input dataset
├── transactions_updated.xlsx  # Output file (generated)
├── requirements.txt        # Dependencies
└── README.md               # Project documentation

🛠️ Technologies Used:

- Python 3.x.
- openpyxl for Excel file processing and chart creation.
- Standard libraries:
    + pathlib for file handling.
    + sys for error management.

📥 Installation

1. Clone the repository:
   - git clone https://github.com/your-username/excel-price-updater.git
   - cd excel-price-updater
2. Create a virtual environment (recommended):
  python -m venv venv
  source venv/bin/activate   # Mac/Linux
  venv\Scripts\activate      # Windows

3. Install dependencies: pip install -r requirements.txt

▶️ Usage
Option 1: Default execution
python main.py
Input file: transactions.xlsx
Output file: transactions_updated.xlsx
Option 2: Command-line version
python main.py input.xlsx output.xlsx

📊 How It Works:

The program loads an Excel workbook. 
It reads the original prices from a specific column.
A correction rule is applied (e.g., 10% discount).
Corrected prices are written into a new column.
A bar chart is generated based on the updated values.
The updated workbook is saved to a new file.

📈 Example Output:

New column: Corrected Price
Bar chart showing updated price distribution
Saved file: transactions_updated.xlsx

👩‍💻 Author : Asmaa CHOUAI - Full-Stack & AI Engineer 

