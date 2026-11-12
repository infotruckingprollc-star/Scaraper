"""
USDOT Scraper for https://motus.dot.gov/public/search

Setup (run once in Command Prompt / Terminal):
    pip install playwright openpyxl
    playwright install chromium

Run:
    python dot_scraper.py
"""

import os
import re
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

SEARCH_URL = "https://motus.dot.gov/public/search"

COLUMNS = [
    "Input Number",
    "USDOT Number",
    "USDOT Status",
    "Legal Business Name",
    "DBA Name",
    "Principal Place of Business",
    "Business Phone",
    "Business Email",
    "Company Officials",
    "Docket Number",
    "Operating Authority Status",
    "Scrape Result",
]

# --------------------------------------------------------------------------
# JavaScript snippets run inside the page
# --------------------------------------------------------------------------

JS_EXTRACT = r"""
() => {
  const clean = s => (s || "").replace(/\s+/g, " ").trim();

  // ---- Header: "USDOT #4552955 -" then "Active"
  let usdot = "", usdotStatus = "";
  const h4s = Array.from(document.querySelectorAll("h4"));
  for (let i = 0; i < h4s.length; i++) {
    const t = clean(h4s[i].textContent);
    const m = t.match(/USDOT\s*#\s*(\d+)/i);
    if (m) {
      usdot = m[1];
      if (h4s[i + 1]) usdotStatus = clean(h4s[i + 1].textContent);
      break;
    }
  }

  // ---- Label/value rows in "Business Information"
  const getField = label => {
    const labels = Array.from(document.querySelectorAll("p.frs-text-bold"));
    for (const p of labels) {
      if (clean(p.textContent).toLowerCase() === label.toLowerCase()) {
        const row = p.parentElement && p.parentElement.parentElement;
        if (row && row.children.length > 1) {
          return clean(row.children[1].textContent);
        }
      }
    }
    return "";
  };

  // ---- Find an accordion by its heading text
  const accordionByTitle = title => {
    const hs = Array.from(document.querySelectorAll("h6"));
    for (const h of hs) {
      if (clean(h.textContent).toLowerCase().startsWith(title.toLowerCase())) {
        return h.closest(".MuiAccordion-root");
      }
    }
    return null;
  };

  // ---- Company officials
  const officials = [];
  const offAcc = accordionByTitle("Company Officials");
  if (offAcc) {
    offAcc.querySelectorAll('[role="row"][data-rowindex]').forEach(r => {
      const name = clean((r.querySelector('[data-field="name"]') || {}).textContent);
      const title = clean((r.querySelector('[data-field="title"]') || {}).textContent);
      if (name) officials.push(title ? name + " (" + title + ")" : name);
    });
  }

  // ---- Operating authority
  const dockets = [], statuses = [];
  const authAcc = accordionByTitle("OPERATING AUTHORITY REGISTRATION");
  if (authAcc) {
    authAcc.querySelectorAll('[role="row"][data-rowindex]').forEach(r => {
      const d = clean((r.querySelector('[data-field="docketNumber"]') || {}).textContent);
      const s = clean((r.querySelector('[data-field="status"]') || {}).textContent);
      if (d) { dockets.push(d); statuses.push(s); }
    });
  }

  return {
    usdot, usdotStatus,
    legalName: getField("Legal Business Name"),
    dba: getField("Doing Business As (DBA) Name"),
    address: getField("Principal Place of Business"),
    phone: getField("Business Telephone No."),
    email: getField("Business Email"),
    officials: officials.join("; "),
    dockets: dockets.join("; "),
    authStatuses: statuses.join("; "),
  };
}
"""

JS_EXPAND_AUTHORITY = r"""
() => {
  const hs = Array.from(document.querySelectorAll("h6"));
  for (const h of hs) {
    if ((h.textContent || "").trim().toUpperCase().startsWith("OPERATING AUTHORITY REGISTRATION")) {
      const btn = h.closest('[role="button"]');
      if (btn && btn.getAttribute("aria-expanded") === "false") btn.click();
      return true;
    }
  }
  return false;
}
"""


# --------------------------------------------------------------------------
# Scraping logic
# --------------------------------------------------------------------------

def parse_numbers(raw):
    """Split pasted text into unique numbers, keeping order."""
    parts = re.split(r"[\s,;]+", raw.strip())
    seen, out = set(), []
    for p in parts:
        digits = re.sub(r"\D", "", p)
        if digits and digits not in seen:
            seen.add(digits)
            out.append(digits)
    return out


def goto_company_page(page, dot):
    """Search for the number on the search page and open the company page.
    Returns True when the company page (with 'USDOT #...') is showing."""
    page.goto(SEARCH_URL, wait_until="domcontentloaded", timeout=60000)

    box = page.locator('input[aria-label^="Search for companies"]')
    box.wait_for(state="visible", timeout=30000)
    box.click()
    box.fill("")
    box.fill(dot)
    page.locator('button[aria-label="search"]').click()

    header = page.locator("h4", has_text=re.compile(r"USDOT\s*#\s*" + dot))
    deadline = time.time() + 25
    clicked = False

    while time.time() < deadline:
        # Already on the company page?
        if header.count() > 0:
            return True

        # On the results list: click the result that contains this number
        if not clicked:
            candidates = [
                page.locator("a", has_text=dot),
                page.locator('[role="row"]', has_text=dot),
                page.locator("tr", has_text=dot),
                page.locator("li", has_text=dot),
                page.get_by_text(dot, exact=False),
            ]
            for c in candidates:
                try:
                    if c.count() > 0 and c.first.is_visible():
                        c.first.click()
                        clicked = True
                        break
                except Exception:
                    pass
        page.wait_for_timeout(700)

    return header.count() > 0


def scrape_one(page, dot):
    row = {c: "" for c in COLUMNS}
    row["Input Number"] = dot

    if not goto_company_page(page, dot):
        row["Scrape Result"] = "Not found / page did not load"
        return row

    # Wait for the business info to render
    try:
        page.locator("p.frs-text-bold", has_text="Legal Business Name").first.wait_for(
            state="attached", timeout=20000
        )
    except PWTimeout:
        pass

    # Make sure the Operating Authority section is expanded so rows render
    try:
        page.evaluate(JS_EXPAND_AUTHORITY)
        page.wait_for_timeout(1200)
    except Exception:
        pass

    d = page.evaluate(JS_EXTRACT)

    row["USDOT Number"] = d["usdot"]
    row["USDOT Status"] = d["usdotStatus"]
    row["Legal Business Name"] = d["legalName"]
    row["DBA Name"] = d["dba"]
    row["Principal Place of Business"] = d["address"]
    row["Business Phone"] = d["phone"]
    row["Business Email"] = d["email"]
    row["Company Officials"] = d["officials"]
    row["Docket Number"] = d["dockets"]
    row["Operating Authority Status"] = d["authStatuses"]
    row["Scrape Result"] = "OK" if d["usdot"] else "Page loaded but no data found"
    return row


# --------------------------------------------------------------------------
# Excel writing
# --------------------------------------------------------------------------

def save_excel(rows, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "DOT Data"

    ws.append(COLUMNS)
    for r in rows:
        ws.append([r[c] for c in COLUMNS])

    head_fill = PatternFill("solid", start_color="1F4E78")
    for cell in ws[1]:
        cell.font = Font(name="Arial", bold=True, color="FFFFFF")
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            # keep numbers like phone / DOT as text so nothing gets mangled
            if isinstance(cell.value, str):
                cell.number_format = "@"

    for i, col in enumerate(COLUMNS, 1):
        longest = max(
            [len(str(col))] + [len(str(r[col])) for r in rows] if rows else [len(col)]
        )
        ws.column_dimensions[get_column_letter(i)].width = min(max(longest + 2, 12), 55)

    ws.freeze_panes = "A2"
    wb.save(path)


# --------------------------------------------------------------------------
# GUI
# --------------------------------------------------------------------------

class App:
    def __init__(self, root):
        self.root = root
        root.title("USDOT Scraper")
        root.geometry("720x640")

        self.stop_flag = threading.Event()
        self.worker = None

        pad = {"padx": 10, "pady": 4}

        tk.Label(root, text="Output file name (without extension):").pack(anchor="w", **pad)
        self.name_var = tk.StringVar(value="dot_results")
        tk.Entry(root, textvariable=self.name_var).pack(fill="x", **pad)

        tk.Label(root, text="Save location:").pack(anchor="w", **pad)
        frame = tk.Frame(root)
        frame.pack(fill="x", **pad)
        self.folder_var = tk.StringVar(value=os.path.expanduser("~"))
        tk.Entry(frame, textvariable=self.folder_var).pack(side="left", fill="x", expand=True)
        tk.Button(frame, text="Browse...", command=self.pick_folder).pack(side="left", padx=(6, 0))

        tk.Label(root, text="Paste DOT numbers (one per line, or separated by commas/spaces):").pack(
            anchor="w", **pad
        )
        self.numbers_box = scrolledtext.ScrolledText(root, height=10)
        self.numbers_box.pack(fill="both", expand=False, **pad)

        opts = tk.Frame(root)
        opts.pack(fill="x", **pad)
        self.show_browser = tk.BooleanVar(value=True)
        tk.Checkbutton(opts, text="Show browser window", variable=self.show_browser).pack(side="left")

        btns = tk.Frame(root)
        btns.pack(fill="x", **pad)
        self.start_btn = tk.Button(
            btns, text="Start Scraping", bg="#1F4E78", fg="white", command=self.start
        )
        self.start_btn.pack(side="left")
        self.stop_btn = tk.Button(btns, text="Stop", command=self.stop, state="disabled")
        self.stop_btn.pack(side="left", padx=6)
        self.progress_var = tk.StringVar(value="Idle")
        tk.Label(btns, textvariable=self.progress_var).pack(side="left", padx=10)

        tk.Label(root, text="Log:").pack(anchor="w", **pad)
        self.log_box = scrolledtext.ScrolledText(root, height=12, state="disabled")
        self.log_box.pack(fill="both", expand=True, **pad)

    # ---- helpers
    def pick_folder(self):
        folder = filedialog.askdirectory(initialdir=self.folder_var.get())
        if folder:
            self.folder_var.set(folder)

    def log(self, msg):
        def _do():
            self.log_box.configure(state="normal")
            self.log_box.insert("end", msg + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")

        self.root.after(0, _do)

    def set_progress(self, text):
        self.root.after(0, lambda: self.progress_var.set(text))

    def set_running(self, running):
        def _do():
            self.start_btn.configure(state="disabled" if running else "normal")
            self.stop_btn.configure(state="normal" if running else "disabled")

        self.root.after(0, _do)

    # ---- actions
    def start(self):
        name = self.name_var.get().strip()
        name = re.sub(r'[\\/:*?"<>|]', "_", name)
        if name.lower().endswith(".xlsx"):
            name = name[:-5]
        folder = self.folder_var.get().strip()
        numbers = parse_numbers(self.numbers_box.get("1.0", "end"))

        if not name:
            messagebox.showerror("Missing", "Please enter a file name.")
            return
        if not os.path.isdir(folder):
            messagebox.showerror("Missing", "Please choose a valid save location.")
            return
        if not numbers:
            messagebox.showerror("Missing", "Please paste at least one DOT number.")
            return

        out_path = os.path.join(folder, name + ".xlsx")
        self.stop_flag.clear()
        self.set_running(True)
        self.log(f"Starting: {len(numbers)} unique numbers -> {out_path}")

        self.worker = threading.Thread(
            target=self.run, args=(numbers, out_path, not self.show_browser.get()), daemon=True
        )
        self.worker.start()

    def stop(self):
        self.stop_flag.set()
        self.log("Stop requested... finishing current number.")

    def run(self, numbers, out_path, headless):
        rows = []
        total = len(numbers)
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=headless)
                context = browser.new_context(viewport={"width": 1366, "height": 900})
                page = context.new_page()
                page.set_default_timeout(30000)

                for i, dot in enumerate(numbers, 1):
                    if self.stop_flag.is_set():
                        break
                    self.set_progress(f"{i}/{total}  (DOT {dot})")
                    try:
                        row = scrape_one(page, dot)
                    except Exception as e:  # keep going on any error
                        row = {c: "" for c in COLUMNS}
                        row["Input Number"] = dot
                        row["Scrape Result"] = f"Error: {str(e)[:150]}"
                    rows.append(row)
                    self.log(f"[{i}/{total}] {dot}: {row['Scrape Result']}  {row['Legal Business Name']}")

                    # save after every number so nothing is lost
                    try:
                        save_excel(rows, out_path)
                    except PermissionError:
                        self.log("!! Could not save - is the Excel file open? Close it; will retry next number.")

                    time.sleep(1.0)

                browser.close()

            save_excel(rows, out_path)
            self.log(f"Done. Saved {len(rows)} rows to {out_path}")
            self.set_progress("Finished")
            self.root.after(0, lambda: messagebox.showinfo("Done", f"Saved to:\n{out_path}"))
        except Exception as e:
            self.log(f"Fatal error: {e}")
            self.set_progress("Error")
            if rows:
                try:
                    save_excel(rows, out_path)
                    self.log(f"Partial results saved to {out_path}")
                except Exception:
                    pass
        finally:
            self.set_running(False)


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
