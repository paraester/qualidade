# qualidade

# How to Use the Quality Control System

## Project Overview
This project automates the generation, manipulation, and filtering of quality data from a web platform. It provides a GUI (Graphical User Interface) for the user to:
- Generate CSV files through a web login and data export.
- Process and filter data based on specific criteria.
- Automatically manipulate names and generate Excel reports for specific individuals.

## Table of Contents
- [Requirements](#requirements)
- [Installation](#installation)
- [Configuration](#configuration)
- [Running the Application](#running-the-application)
- [Main Features](#main-features)
- [File Structure](#file-structure)
- [Notes](#notes)

## Requirements
Before running this project, ensure that the following software and libraries are installed:

1. **Python 3.x**
2. **pip** (Python package installer)
3. **Required Python Libraries**:
   - `pandas`
   - `openpyxl`
   - `tkinter`
   - `ttkbootstrap`
   - `selenium`
   - `pyautogui`
   - `webdriver_manager`

## Installation

1. Clone the project repository:

   ```bash
   git clone <your-repository-url>
   cd <your-repository-folder>
   ```

2. Install the required Python packages:

   ```bash
   pip install -r requirements.txt
   ```

   If the `requirements.txt` does not exist, you can manually install the dependencies with:

   ```bash
   pip install pandas openpyxl selenium pyautogui ttkbootstrap webdriver_manager
   ```

## Configuration

1. **Credentials Setup**:
   
   The project uses a `credentials.py` file for storing the web platform login credentials. Make sure this file exists and contains the following structure:
   
   ```python
   # credentials.py
   username = "your_username"
   password = "your_password"
   ```

2. **WebDriver Configuration**:
   
   The Selenium WebDriver uses Firefox by default. Ensure Firefox is installed on your system. The `webdriver_manager` package will automatically manage the browser drivers.

3. **Excel Template**:
   
   The project uses an Excel template `modelo.xlsx` to generate the final reports with individual sheets for each person. Make sure the `modelo.xlsx` file is placed in the project root directory.

## Running the Application

To start the application and access the GUI:

1. Navigate to the project folder:

   ```bash
   cd <your-repository-folder>
   ```

2. Run the `interface_tk-2.py` script:

   ```bash
   python3 interface_tk-2.py
   ```

   This will launch a GUI where you can generate reports and manipulate data.

## Main Features

### 1. **Generate Quality Data**
- From the GUI, click the **"Gerar arquivo Qualidade"** button to automatically log in to the web platform and download the CSV file.

### 2. **Process and Filter CSV**
- After generating the CSV, you can process the file by clicking the **"Tratar arquivo Qualidade"** button.
- You can also filter the data by selecting a responsible area in the **Combobox** and clicking **"Gerar Arquivo Filtrado"**.

### 3. **Name Manipulation**
- The **"Inverter Nomes"** button allows you to reverse names in the format `Last Name, First Name` to `First Name Last Name`, and the file will be saved automatically.

### 4. **Generate Excel Reports**
- The **"Inverter e Criar Abas no Excel"** function processes the data, creating individual Excel sheets for each person listed under the "Nome completo" column, using a pre-designed template (`modelo.xlsx`).

## File Structure

```bash
.
├── Dados/                     # Folder where the generated CSVs and reports are stored
├── interface_tk-2.py           # Main GUI script
├── login_e_exportar.py         # Script to log in and export CSV
├── tratar.py                   # Script to manipulate and clean CSV
├── modelo.xlsx                 # Excel template for generating reports
├── credentials.py              # File containing username and password
└── requirements.txt            # File with required Python dependencies
```

## Notes

- The **webdriver_manager** automatically manages the required drivers for Selenium.
- The **`modelo.xlsx`** must contain a sheet named `'primeiroNome'`, which acts as the template for individual reports.
- The GUI uses the **ttkbootstrap** theme (`cosmo`), providing a modern look and feel for the user interface.
