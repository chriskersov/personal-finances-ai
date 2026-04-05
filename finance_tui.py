import os
import re
import ollama
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string as col_idx

from textual.app import App, ComposeResult
from textual.containers import ScrollableContainer, Horizontal, Vertical
from textual.widgets import Footer, Static, Button, ListView, ListItem, Label
from textual.worker import Worker, WorkerState
from textual_plotext import PlotextPlot

# ── Config ───────────────────────────────────────────────────────────────────
MODEL = "qwen2.5:7b"
EXCEL_PATH = os.environ.get("FINANCE_FILE", "finances.xlsm")

# Only show sheets that match "Month Year" e.g. "April 2025"
MONTH_SHEET_RE = re.compile(
    r"^(January|February|March|April|May|June|July|August|September|October|November|December)"
    r"\s+\d{4}$"
)

# ── Pure Logic ────────────────────────────────────────────────────────────────
def c(ws, col_letter, row):
    v = ws.cell(row, col_idx(col_letter)).value
    if v is None or v == "" or str(v) == "#DIV/0!":
        return None
    return v

def fmt(v):
    if v is None: return "—"
    try: return f"£{float(v):,.2f}"
    except: return str(v)

def fmt_pct(v):
    if v is None: return "—"
    try: return f"{float(v)*100:.0f}%"
    except: return str(v)


def load_sheet(path, sheet_name):
    wb = load_workbook(path, data_only=True)
    ws = wb[sheet_name]

    needs = []
    for r in range(24, 101):
        desc, date, cat, actual = c(ws,'B',r), c(ws,'C',r), c(ws,'D',r), c(ws,'E',r)
        if actual is not None:
            needs.append({"description": desc, "date": str(date)[:10] if date else None,
                          "category": cat, "actual": actual})

    wants = []
    for r in range(24, 101):
        desc, date, cat, actual = c(ws,'G',r), c(ws,'H',r), c(ws,'I',r), c(ws,'J',r)
        if actual is not None:
            wants.append({"description": desc, "date": str(date)[:10] if date else None,
                          "category": cat, "actual": actual})

    savings = []
    for r in range(24, 101):
        desc, date, cat, actual = c(ws,'L',r), c(ws,'M',r), c(ws,'N',r), c(ws,'O',r)
        if actual is not None:
            savings.append({"description": desc, "date": str(date)[:10] if date else None,
                            "category": cat, "actual": actual})

    cash_flow = {}
    for row, label in [(43,"needs"),(44,"wants"),(45,"savings"),(46,"total")]:
        cash_flow[label] = {
            "in":   c(ws,'U',row),
            "out":  c(ws,'V',row),
            "diff": c(ws,'W',row),
        }

    goals = {}
    for row, label in [(51,"needs"),(52,"wants"),(53,"savings")]:
        goals[label] = {
            "goal":   c(ws,'T',row),
            "actual": c(ws,'V',row),
        }

    income = []
    for r in range(58, 101):
        desc, date, actual = c(ws,'Q',r), c(ws,'T',r), c(ws,'V',r)
        if actual is not None:
            income.append({"description": desc,
                           "date": str(date)[:10] if date else None,
                           "actual": actual})

    categories = []
    for r in range(53, 67):
        name = c(ws,'AA',r)
        if name:
            categories.append({
                "category": name,
                "needs":    c(ws,'AB',r),
                "wants":    c(ws,'AC',r),
                "savings":  c(ws,'AD',r),
                "total":    c(ws,'AE',r),
            })

    return {
        "needs": needs, "wants": wants, "savings": savings,
        "cash_flow": cash_flow, "goals": goals,
        "income": income, "categories": categories,
    }


def build_context(sheet_name, d):
    lines = []
    lines.append(f"MONTH: {sheet_name}")
    lines.append(f"Data read on: {datetime.now().strftime('%d %B %Y')}\n")

    cf = d["cash_flow"]
    lines.append("CASH FLOW:")
    for label in ["needs","wants","savings","total"]:
        v = cf[label]
        lines.append(f"  {label.capitalize():8s} — In: {fmt(v['in'])}  Out: {fmt(v['out'])}  Difference: {fmt(v['diff'])}")

    lines.append("\nBUDGET GOALS vs ACTUAL:")
    for label, v in d["goals"].items():
        lines.append(f"  {label.capitalize():8s} — Goal: {fmt_pct(v['goal'])}  Actual: {fmt_pct(v['actual'])}")

    if d["income"]:
        lines.append(f"\nINCOME ({len(d['income'])} entries):")
        for e in d["income"]:
            lines.append(f"  {e['date']} | {e['description']} | {fmt(e['actual'])}")
    else:
        lines.append("\nINCOME: none recorded yet")

    if d["needs"]:
        total = sum(float(e["actual"]) for e in d["needs"])
        lines.append(f"\nNEEDS SPENDING — {len(d['needs'])} entries, total {fmt(total)}:")
        for e in d["needs"]:
            lines.append(f"  {e['date']} | {e['category']} | {e['description']} | {fmt(e['actual'])}")
    else:
        lines.append("\nNEEDS SPENDING: none recorded yet")

    if d["wants"]:
        total = sum(float(e["actual"]) for e in d["wants"])
        lines.append(f"\nWANTS SPENDING — {len(d['wants'])} entries, total {fmt(total)}:")
        for e in d["wants"]:
            lines.append(f"  {e['date']} | {e['category']} | {e['description']} | {fmt(e['actual'])}")
    else:
        lines.append("\nWANTS SPENDING: none recorded yet")

    if d["savings"]:
        total = sum(float(e["actual"]) for e in d["savings"])
        lines.append(f"\nSAVINGS — {len(d['savings'])} entries, total {fmt(total)}:")
        for e in d["savings"]:
            lines.append(f"  {e['date']} | {e['category']} | {e['description']} | {fmt(e['actual'])}")
    else:
        lines.append("\nSAVINGS: none recorded yet")

    active_cats = [cat for cat in d["categories"] if cat["total"] not in (None, 0)]
    if active_cats:
        lines.append(f"\nSPEND BY CATEGORY:")
        for cat in sorted(active_cats, key=lambda x: float(x["total"] or 0), reverse=True):
            lines.append(f"  {cat['category']:15s} — Needs: {fmt(cat['needs'])}  Wants: {fmt(cat['wants'])}  Total: {fmt(cat['total'])}")
    else:
        lines.append("\nSPEND BY CATEGORY: no data yet")

    return "\n".join(lines)


def generate_digest(context, sheet_name):
    now = datetime.now()
    day_of_month = now.day
    days_in_month = 31 if now.month in [1,3,5,7,8,10,12] else 30 if now.month in [4,6,9,11] else 28
    pct_through = round((day_of_month / days_in_month) * 100)

    prompt = f"""Write a monthly financial summary for {sheet_name}.
Today is {now.strftime('%d %B %Y')} — day {day_of_month} of {days_in_month} ({pct_through}% through the month).

Use only the numbers in the data. Be specific. Write 4-5 sentences covering:
- Total income vs total spending and the net position
- How the needs / wants / savings split compares to the goals (25% / 15% / 60%)
- The biggest spending category
- One honest, useful observation for the rest of the month

If any section shows no data or zeros, say so clearly.
Write as a plain paragraph. No bullet points. Calm, direct tone like a smart financial advisor."""

    response = ollama.chat(
        model=MODEL,
        messages=[
            {"role": "system", "content": f"You are a personal finance assistant. Here is the user's full finance data for this month:\n\n{context}"},
            {"role": "user", "content": prompt}
        ]
    )
    return response["message"]["content"]


# ── TUI Application ──────────────────────────────────────────────────────────
class FinanceApp(App):
    CSS = """
    Screen { background: #101112; color: #e8e4dc; }

    #title {
        color: #5fffb0;
        text-style: bold;
        margin: 1 0 0 2;
    }

    #divider {
        color: #2b3330;
        margin: 0 2;
    }

    #content-row {
        margin: 0 2;
        height: 1fr;
    }

    #sidebar {
        width: 24;
        border-right: solid #2b3330;
        padding: 0 1 0 0;
    }

    #sidebar-title {
        color: #5fffb0;
        text-style: bold;
        margin: 1 0 1 0;
    }

    ListView {
        background: #101112;
        border: none;
        padding: 0;
    }

    ListItem {
        background: #101112;
        color: #6f736f;
        padding: 0 2;
    }

    ListItem:hover {
        background: #1a1f1d;
        color: #d6d2c8;
    }

    ListItem.-highlight {
        background: #1d2421;
        color: #5fffb0;
    }

    ListItem.current-month {
        color: #c8c4bc;
    }

    ListItem.selected-month {
        background: #1d2421;
        color: #5fffb0;
        text-style: bold;
    }

    #main {
        padding: 0;
    }

    #selected-month {
        margin: 1 0 0 2;
        text-style: bold;
        color: #e8e4dc;
    }

    #digest-box {
        background: #101112;
        border: solid #2b3330;
        border-left: tall #5fffb0;
        padding: 1 2;
        margin: 1 2 0 2;
        height: auto;
        color: #d6d2c8;
    }

    #digest-box.loading {
        color: #5fffb0;
        text-style: italic;
    }

    #digest-box.error-box {
        background: #180f0f;
        border: solid #3a1a1a;
        border-left: tall #ff5f5f;
        color: #ff9a9a;
    }

    #chart {
        height: 20;
        margin: 1 2 0 2;
        border: solid #2b3330;
        border-left: tall #5fffb0;
        background: #101112;
    }

    #regen {
        background: #171c1a;
        color: #5fffb0;
        border: solid #2b3330;
        margin-left: 2;
        margin-top: 1;
        margin-bottom: 1;
    }

    #regen:hover {
        background: #24302b;
        color: #5fffb0;
        border: solid #5fffb0;
    }

    #regen:disabled {
        background: #101112;
        color: #435049;
        border: solid #222826;
    }

    Footer {
        background: #171c1a;
        color: #5fffb0;
    }
    """

    def __init__(self):
        super().__init__()
        self._sheet_names: list[str] = []
        self._selected_sheet: str = ""

    def compose(self) -> ComposeResult:
        yield Static("FINANCE AI", id="title")
        yield Static("─" * 80, id="divider")
        with Horizontal(id="content-row"):
            with Vertical(id="sidebar"):
                yield Static("MONTHS", id="sidebar-title")
                yield ListView(id="sheet-list")
            with ScrollableContainer(id="main"):
                yield Static("", id="selected-month")
                yield Static("Select a month to begin…", id="digest-box")
                yield PlotextPlot(id="chart")
                yield Button("↻  REGENERATE", id="regen")
        yield Footer()

    def on_mount(self) -> None:
        # Hide chart until data loads
        self.query_one("#chart").display = False

        if not os.path.exists(EXCEL_PATH):
            self._show_error(
                f"Excel file not found at: {EXCEL_PATH}\n\n"
                f"Set the path with:\n  export FINANCE_FILE=/full/path/to/your/file.xlsm"
            )
            return
        self._load_sheets()

    def _load_sheets(self) -> None:
        wb = load_workbook(EXCEL_PATH, read_only=True)
        all_sheets = wb.sheetnames
        wb.close()

        self._sheet_names = [n for n in all_sheets if MONTH_SHEET_RE.match(n)]

        now = datetime.now()
        current_month_str = now.strftime("%B %Y")

        list_view = self.query_one("#sheet-list", ListView)

        for name in self._sheet_names:
            is_current = current_month_str.lower() in name.lower()
            item = ListItem(Label(name), name=name)
            if is_current:
                item.add_class("current-month")
            list_view.append(item)

        list_view.index = next(
            (i for i, n in enumerate(self._sheet_names)
             if current_month_str.lower() in n.lower()),
            0
        )

        if list_view.children:
            current_item = list_view.children[list_view.index]
            if isinstance(current_item, ListItem):
                self._set_selected_item_class(current_item)

    def _set_selected_item_class(self, selected_item: ListItem) -> None:
        for item in self.query("#sheet-list ListItem"):
            item.remove_class("selected-month")
        selected_item.add_class("selected-month")

    def _update_chart(self, categories: list) -> None:
        active = [cat for cat in categories if cat["total"] not in (None, 0)]
        active = sorted(active, key=lambda x: float(x["total"] or 0), reverse=True)

        chart_widget = self.query_one("#chart", PlotextPlot)

        if not active:
            chart_widget.display = False
            return

        plt = chart_widget.plt
        plt.clear_figure()
        plt.theme("dark")

        labels = [cat["category"] for cat in active]
        needs_vals  = [float(cat["needs"]  or 0) for cat in active]
        wants_vals  = [float(cat["wants"]  or 0) for cat in active]

        plt.stacked_bar(labels, [needs_vals, wants_vals], labels=["Needs", "Wants"], color=["green", "cyan"])
        plt.title("Spend by Category")
        plt.xlabel("")
        plt.yfrequency(4)

        chart_widget.display = True
        chart_widget.refresh()

    def on_list_view_selected(self, event: ListView.Selected) -> None:
        sheet_name = event.item.name or ""
        if sheet_name in self._sheet_names:
            self._set_selected_item_class(event.item)
            self._selected_sheet = sheet_name
            self.query_one("#selected-month", Static).update(sheet_name)
            self._run_digest()

    def _run_digest(self) -> None:
        if not self._selected_sheet:
            return
        sheet = self._selected_sheet
        box = self.query_one("#digest-box", Static)
        box.remove_class("error-box")
        box.update("Reading your finances…")
        box.add_class("loading")
        self.query_one("#regen", Button).disabled = True
        self.query_one("#chart").display = False
        self.run_worker(
            lambda: self._worker_logic(sheet),
            thread=True,
            exclusive=True,
        )

    def _worker_logic(self, sheet: str) -> dict:
        data = load_sheet(EXCEL_PATH, sheet)
        context = build_context(sheet, data)
        digest = generate_digest(context, sheet)
        return {"digest": digest, "categories": data["categories"]}

    def on_worker_state_changed(self, event: Worker.StateChanged) -> None:
        box = self.query_one("#digest-box", Static)
        btn = self.query_one("#regen", Button)

        if event.state == WorkerState.SUCCESS:
            result = event.worker.result
            box.remove_class("loading")
            box.update(result["digest"])
            btn.disabled = False
            self._update_chart(result["categories"])

        elif event.state == WorkerState.ERROR:
            box.remove_class("loading")
            box.add_class("error-box")
            box.update(f"Error: {event.worker.error}")
            btn.disabled = False

    def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "regen":
            self._run_digest()

    def _show_error(self, message: str) -> None:
        box = self.query_one("#digest-box", Static)
        box.remove_class("loading")
        box.add_class("error-box")
        box.update(message)


if __name__ == "__main__":
    FinanceApp().run()