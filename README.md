# Job Application Bot

This project automates the process of finding job opportunities, generating application materials, and saving them in structured formats (Excel, Word, and text).

## ✨ Features
- 🔎 Searches for job opportunities via **JSearch API** (RapidAPI).
- 📑 Saves all opportunities in an **Excel file** with company, title, location, requirements, and link.
- 📧 Generates **email drafts** for each job.
- 📝 Generates **motivation letters (Word `.docx`)** for each job.
- ⚡ Can be extended with Jobbörse API integration.

## 📂 Output
All generated files (Excel, emails, motivation letters) are saved to the `application_bot_outputs/` directory.

## 🛠 Requirements
- Python 3.8+
- Dependencies:
  - `requests`
  - `python-docx`
  - `openpyxl`

Install them with:
```bash
pip install requests python-docx openpyxl
