"""
Simple Database Handling System (no SQL)
- Data is stored in a CSV file (records.csv)
- Interface is a Tkinter form (no console)
- Features: Add, Update, Delete (by ID dialog), right-click Delete/Edit,
  live Search across every field, Show All, Export to Excel

To change the fields, just edit the FIELDS list below.
The first field is treated as the unique key (ID).
"""

import csv
import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

# Always save next to this program file, whatever folder it is run from
FILE_NAME = os.path.join(os.path.dirname(os.path.abspath(__file__)), "records.csv")
FIELDS = ["ID", "Name", "Class", "Phone", "Email"]   # <-- edit these
KEY = FIELDS[0]

# ---------- Colour theme (change these to restyle the whole app) ----------
BG      = "#f4f6f8"   # window background
HEADER  = "#1f3a5f"   # navy banner + table heading
ACCENT  = "#2e86de"   # blue (update, search, selection)
SUCCESS = "#27ae60"   # green (add)
DANGER  = "#d63031"   # red (delete)
NEUTRAL = "#636e72"   # grey (clear, show all, export)
TEXT    = "#2d3436"   # normal text
ROW_ALT = "#eaf2fb"   # alternate table row
FONT    = ("Segoe UI", 10)


# ---------- File handling (the "database" part) ----------
def load_records():
    """Read all rows from the CSV file. Creates the file if missing."""
    if not os.path.exists(FILE_NAME):
        with open(FILE_NAME, "w", newline="", encoding="utf-8") as f:
            csv.DictWriter(f, fieldnames=FIELDS).writeheader()
        return []
    with open(FILE_NAME, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def save_records(rows):
    """Write all rows back to the CSV file."""
    with open(FILE_NAME, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDS)
        writer.writeheader()
        writer.writerows(rows)


# ---------- GUI ----------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Database Handling System  -  saving to: " + FILE_NAME)
        self.geometry("900x580")
        self.configure(bg=BG)
        self.entries = {}
        self.apply_theme()
        self.build_header()
        self.build_form()
        self.build_buttons()
        self.build_table()
        self.build_context_menu()
        self.refresh_table(load_records())

    def apply_theme(self):
        style = ttk.Style(self)
        style.theme_use("clam")          # 'clam' lets buttons/tables take colours
        style.configure(".", background=BG, foreground=TEXT, font=FONT)
        style.configure("TLabelframe", background=BG, bordercolor="#c8d0d6")
        style.configure("TLabelframe.Label", background=BG, foreground=HEADER,
                        font=("Segoe UI", 11, "bold"))
        style.configure("TEntry", fieldbackground="white", padding=4)

        # coloured buttons with hover effect
        for name, colour, hover in [("Success", SUCCESS, "#1e8e4e"),
                                    ("Accent", ACCENT, "#1b6ec2"),
                                    ("Danger", DANGER, "#b02526"),
                                    ("Neutral", NEUTRAL, "#4b5457")]:
            style.configure(f"{name}.TButton", background=colour, foreground="white",
                            font=("Segoe UI", 10, "bold"), padding=(12, 6),
                            borderwidth=0, focuscolor=colour)
            style.map(f"{name}.TButton", background=[("active", hover)],
                      foreground=[("active", "white")])

        # table
        style.configure("Treeview", background="white", fieldbackground="white",
                        foreground=TEXT, rowheight=28, bordercolor="#c8d0d6")
        style.configure("Treeview.Heading", background=HEADER, foreground="white",
                        font=("Segoe UI", 10, "bold"), padding=6, relief="flat")
        style.map("Treeview.Heading", background=[("active", ACCENT)])
        style.map("Treeview", background=[("selected", ACCENT)],
                  foreground=[("selected", "white")])

    def build_header(self):
        banner = tk.Frame(self, bg=HEADER)
        banner.pack(fill="x")
        tk.Label(banner, text="Database Handling System", bg=HEADER, fg="white",
                 font=("Segoe UI", 18, "bold")).pack(side="left", padx=16, pady=(10, 0))
        tk.Label(banner, text="Add  |  Update  |  Delete  |  Search", bg=HEADER,
                 fg="#b8c7dc", font=("Segoe UI", 10)).pack(side="left", padx=4, pady=(16, 0))
        tk.Frame(self, bg=ACCENT, height=4).pack(fill="x")   # accent stripe

    def build_form(self):
        frame = ttk.LabelFrame(self, text="Entry Details")
        frame.pack(fill="x", padx=10, pady=8)
        for i, field in enumerate(FIELDS):
            ttk.Label(frame, text=field).grid(row=i // 2, column=(i % 2) * 2,
                                              padx=8, pady=6, sticky="e")
            entry = ttk.Entry(frame, width=30)
            entry.grid(row=i // 2, column=(i % 2) * 2 + 1, padx=8, pady=6)
            self.entries[field] = entry

    def build_buttons(self):
        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=10)
        for text, cmd, colour in [("Add", self.add, "Success"),
                                  ("Update", self.update_record, "Accent"),
                                  ("Delete", self.delete_by_dialog, "Danger"),
                                  ("Clear", self.clear_form, "Neutral"),
                                  ("Show All", self.show_all, "Neutral"),
                                  ("Export to Excel", self.export_excel, "Neutral")]:
            ttk.Button(bar, text=text, command=cmd,
                       style=f"{colour}.TButton").pack(side="left", padx=4, pady=6)

        ttk.Label(bar, text="Search:").pack(side="left", padx=(20, 2))
        self.search_var = tk.StringVar()
        search = ttk.Entry(bar, textvariable=self.search_var, width=22)
        search.pack(side="left")
        # live search: results update as you type
        search.bind("<KeyRelease>", lambda e: self.search(silent=True))
        search.bind("<Return>", lambda e: self.search())
        ttk.Button(bar, text="Go", command=self.search, style="Accent.TButton").pack(side="left", padx=4)

    def build_table(self):
        frame = ttk.Frame(self)
        frame.pack(fill="both", expand=True, padx=10, pady=8)
        self.table = ttk.Treeview(frame, columns=FIELDS, show="headings")
        for f in FIELDS:
            self.table.heading(f, text=f)
            self.table.column(f, width=120)
        scroll = ttk.Scrollbar(frame, orient="vertical", command=self.table.yview)
        self.table.configure(yscrollcommand=scroll.set)
        self.table.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.table.bind("<<TreeviewSelect>>", self.on_select)

    def build_context_menu(self):
        """Right-click menu on each entry."""
        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Edit (load into form)", command=self.on_select)
        self.menu.add_command(label="Delete this record", command=self.delete_selected)
        for button in ("<Button-3>", "<Button-2>", "<Control-Button-1>"):
            self.table.bind(button, self.show_menu)

    def show_menu(self, event):
        row = self.table.identify_row(event.y)
        if not row:
            return
        self.table.selection_set(row)
        try:
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu.grab_release()

    # ----- helpers -----
    def get_form(self):
        return {f: self.entries[f].get().strip() for f in FIELDS}

    def clear_form(self):
        for e in self.entries.values():
            e.delete(0, tk.END)

    def refresh_table(self, rows):
        self.table.delete(*self.table.get_children())
        self.table.tag_configure("odd", background=ROW_ALT)
        self.table.tag_configure("even", background="white")
        for i, r in enumerate(rows):
            # the row id (iid) is the record's ID, so values never get altered
            self.table.insert("", tk.END, iid=r[KEY], values=[r[f] for f in FIELDS],
                              tags=("odd" if i % 2 else "even",))

    def on_select(self, _event=None):
        sel = self.table.selection()
        if not sel:
            return
        key_value = sel[0]
        for r in load_records():
            if r[KEY] == key_value:
                self.clear_form()
                for f in FIELDS:
                    self.entries[f].insert(0, r[f])
                return

    def remove_record(self, key_value):
        """Ask for confirmation, then delete the record with this ID."""
        rows = load_records()
        new_rows = [r for r in rows if r[KEY] != key_value]
        if len(new_rows) == len(rows):
            messagebox.showerror("Not found", f"No record with {KEY} '{key_value}'.")
            return
        if messagebox.askyesno("Confirm", f"Delete record with {KEY} '{key_value}'?"):
            save_records(new_rows)
            self.search(silent=True)
            self.clear_form()

    # ----- actions -----
    def add(self):
        record = self.get_form()
        if not record[KEY]:
            messagebox.showwarning("Missing", f"{KEY} is required.")
            return
        rows = load_records()
        if any(r[KEY] == record[KEY] for r in rows):
            messagebox.showerror("Duplicate", f"A record with this {KEY} already exists.")
            return
        rows.append(record)
        save_records(rows)
        self.refresh_all()
        self.clear_form()
        messagebox.showinfo("Saved", "Record added.")

    def update_record(self):
        record = self.get_form()
        rows = load_records()
        for i, r in enumerate(rows):
            if r[KEY] == record[KEY]:
                rows[i] = record
                save_records(rows)
                self.search(silent=True)
                messagebox.showinfo("Updated", "Record updated.")
                return
        messagebox.showerror("Not found", f"No record with that {KEY}.")

    def delete_by_dialog(self):
        """Regular Delete button: asks for the ID in its own dialog box."""
        value = simpledialog.askstring("Delete Record",
                                       f"Enter the {KEY} of the record to delete:",
                                       parent=self)
        if value is None:          # Cancel pressed
            return
        value = value.strip()
        if not value:
            messagebox.showwarning("Missing", f"Please enter a {KEY}.")
            return
        self.remove_record(value)

    def delete_selected(self):
        """Right-click Delete: deletes the entry that was clicked."""
        sel = self.table.selection()
        if sel:
            self.remove_record(sel[0])

    def search(self, silent=False):
        """Matches ANY text in ANY field. Several words = all must match."""
        words = self.search_var.get().lower().split()
        rows = load_records()
        found = [r for r in rows
                 if all(w in " ".join(str(v) for v in r.values()).lower() for w in words)]
        self.refresh_table(found)
        if not found and not silent:
            messagebox.showinfo("Search", "No matching records.")

    def refresh_all(self):
        """Clear the search box and list every record in the table."""
        self.search_var.set("")
        self.refresh_table(load_records())

    def show_all(self):
        """Show All button: list every record AND open records.csv."""
        self.refresh_all()
        self.open_csv()

    def open_csv(self):
        """Open records.csv in the default program (usually Excel)."""
        load_records()   # makes sure the file exists
        try:
            if sys.platform.startswith("win"):
                os.startfile(FILE_NAME)
            elif sys.platform == "darwin":
                subprocess.call(["open", FILE_NAME])
            else:
                subprocess.call(["xdg-open", FILE_NAME])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file:\n{FILE_NAME}\n\n{e}")

    def export_excel(self):
        try:
            from openpyxl import Workbook
        except ImportError:
            messagebox.showerror("Missing library",
                                 "Install it first:  pip install openpyxl")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel file", "*.xlsx")])
        if not path:
            return
        wb = Workbook()
        ws = wb.active
        ws.append(FIELDS)
        for r in load_records():
            ws.append([r[f] for f in FIELDS])
        wb.save(path)
        messagebox.showinfo("Exported", "Data exported to Excel.")


if __name__ == "__main__":
    App().mainloop()
