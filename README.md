<!-- ```
 ____  _____ ____  ____   ___  _   _    _    _       _____ ___ _   _    _    _   _  ____ _____ ____      _    ___
|  _ \| ____|  _ \/ ___| / _ \| \ | |  / \  | |     |  ___|_ _| \ | |  / \  | \ | |/ ___| ____/ ___|    / \  |_ _|
| |_) |  _| | |_) \___ \| | | |  \| | / _ \ | |     | |_   | ||  \| | / _ \ |  \| | |   |  _| \___ \   / _ \  | |
|  __/| |___|  _ < ___) | |_| | |\  |/ ___ \| |___  |  _|  | || |\  |/ ___ \| |\  | |___| |___ ___) | / ___ \ | |
|_|   |_____|_| \_\____/ \___/|_| \_/_/   \_\_____| |_|   |___|_| \_/_/   \_\_| \_|\____|_____|____/ /_/   \_\___|
``` -->

![Personal Finances AI](./title.svg)

> `// personal finance tracker with a local AI layer`

---

## `WHAT IS THIS`

I keep my finances in an Excel workbook - one sheet per month, tracking needs, wants, savings, income, and spending by category. This is a Streamlit app that reads that workbook and uses a locally running LLM to generate a plain-English summary of the current month. Everything runs on my machine, the Excel file is never committed to this repo.

---

## `STACK`

|           |                                             |
| --------- | ------------------------------------------- |
| interface | [Streamlit](https://streamlit.io)           |
| data      | [openpyxl](https://openpyxl.readthedocs.io) |
| local LLM | [Ollama](https://ollama.com) - `qwen2.5:7b` |

---

## `SETUP`

```bash
git clone https://github.com/YOUR_USERNAME/personal-finances-ai
cd personal-finances-ai
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
ollama pull qwen2.5:7b
```

```bash
export FINANCE_FILE=/path/to/your/finances.xlsm
streamlit run finance_app.py
```

---

## `CURRENT STATE`

```
v0.1  ████████░░░░░░░░░░░░  in progress
```

Right now it does one thing - generates a monthly digest from the current sheet. A short paragraph covering cash flow, budget goal tracking, top spending categories, and one observation for the rest of the month.

A lot more is planned: chat interface, multi-month comparisons, budget alerts, spend forecasting, and more.

---

`// runs fully offline · Excel file is gitignored · nothing leaves the machine`
