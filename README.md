# 🎓 bputResultsVance

A Python-based BPUT results scraper that retrieves a student's academic results using their registration number and exports them into a formatted Excel workbook (`.xlsx`).

**BPUT Results Portal:** https://results.bput.ac.in/

## ✨ Features

- 📚 Retrieves available semester-wise academic results.
- ⚡ Concurrent examination-session scanning.
- 📊 Calculates CGPA, credits, grade points, and backlog status.
- 🖥️ Interactive terminal interface with progress indicators.
- 📗 Generates a formatted Excel workbook with three sheets:
  - **Academic Overview** — Student profile and semester summaries.
  - **Consolidated Grade Sheet** — Detailed semester-wise results.
  - **All Subjects Master List** — Filterable list of subjects.
- 🎨 Color-coded grades and formatted tables.

## 🚀 Installation

### 1. Download the Repository

Clone the repository using Git:

```bash
git clone https://github.com/Bibek-M/BputResultsVance.git
cd bputResultsVance
```

Alternatively, download the repository as a ZIP from GitHub and extract it.

### 2. Install Python

Check whether Python is installed:

```bash
python --version
```

If that doesn't work, try `python3 --version` or, on Windows, `py --version`.

If Python is not installed, download it from the official website:

**https://www.python.org/downloads/**

After installation, reopen your terminal and verify the version.

### 3. Install Dependencies

Run the appropriate command from the project directory.

**Windows:**

```bash
python -m pip install -r requirements.txt

or

run the start.bat directly to install and run
```
          
          

**macOS / Linux:**

```bash
python3 -m pip install -r requirements.txt
```

## ▶️ Usage

### Interactive Mode

**Windows:**

```bash
python main.py
```

**macOS / Linux:**

```bash
python3 main.py
```

Follow the terminal prompts to enter the student's registration number and choose the examination sessions to retrieve.

### Command-Line Options

```bash
# Retrieve all available semesters
python main.py --reg YOUR_REGISTRATION_NUMBER --all

# Retrieve results and open the Excel workbook
python main.py --reg YOUR_REGISTRATION_NUMBER --all --open

# Retrieve a specific examination session
python main.py --reg YOUR_REGISTRATION_NUMBER --session "Even-(2023-24)"

# Specify the output file
python main.py --reg YOUR_REGISTRATION_NUMBER --all -o "MyResult.xlsx"
```

Replace `YOUR_REGISTRATION_NUMBER` with the student's registration number. On macOS/Linux, use `python3` instead of `python` where necessary.

## 📁 Project Structure

```text
bputResultsVance/
├── main.py             # Application entry point
├── cli.py              # Terminal interface
├── bput_scraper.py     # Results retrieval and session scanning
├── excel_exporter.py   # Excel workbook generation
├── requirements.txt    # Python dependencies
├── README.md           # Project documentation
└── LICENSE             # MIT License
```

## ⚠️ Disclaimer

- This is an independent project and is not affiliated with or endorsed by BPUT.
- Results retrieval depends on the availability and behavior of the official portal.
- Verify calculated academic summaries against official university records.
- Use the scraper responsibly and avoid excessive requests.

## 📜 License

This project is licensed under the **MIT License**. See the [LICENSE](LICENSE) file for details.

---

**Made for BPUT students. 🎓**