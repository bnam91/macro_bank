import threading
import tkinter as tk
from tkinter import ttk, messagebox

from config import SPREADSHEET_ID
from utils.sheets import fetch_sheet_names, write_name_rrn


def load_sheet_names(spreadsheet_id):
    try:
        return [
            name
            for name in fetch_sheet_names(spreadsheet_id)
            if "완료_" not in name
        ]
    except Exception as e:
        messagebox.showerror("오류", f"시트 목록 불러오기 실패: {e}")
        return []


def run():
    spreadsheet_id = SPREADSHEET_ID

    root = tk.Tk()
    root.title("입금요청봇")
    root.geometry("420x220")
    root.resizable(False, False)

    frame = ttk.Frame(root, padding=16)
    frame.pack(fill=tk.BOTH, expand=True)

    name_var = tk.StringVar()
    rrn_var = tk.StringVar()
    sheet_var = tk.StringVar()
    all_sheet_names = []

    ttk.Label(frame, text="시트").grid(row=0, column=0, sticky="w")
    sheet_combo = ttk.Combobox(
        frame, textvariable=sheet_var, width=28, height=12, state="readonly"
    )
    sheet_combo.grid(row=0, column=1, sticky="ew", pady=4)

    ttk.Label(frame, text="이름").grid(row=1, column=0, sticky="w")
    name_entry = ttk.Entry(frame, textvariable=name_var, width=30)
    name_entry.grid(row=1, column=1, sticky="ew", pady=4)

    ttk.Label(frame, text="주민번호").grid(row=2, column=0, sticky="w")
    rrn_entry = ttk.Entry(frame, textvariable=rrn_var, width=30)
    rrn_entry.grid(row=2, column=1, sticky="ew", pady=4)

    def set_combo_values(values):
        sheet_combo["values"] = values
        if values:
            sheet_combo.current(0)
        if not values:
            sheet_var.set("")

    def refresh_sheets():
        def worker():
            names = load_sheet_names(spreadsheet_id)

            def apply_result():
                nonlocal all_sheet_names
                all_sheet_names = names
                set_combo_values(all_sheet_names)
                refresh_btn.config(state="normal")
                sheet_combo.config(state="normal")

            root.after(0, apply_result)

        refresh_btn.config(state="disabled")
        sheet_combo.config(state="disabled")
        threading.Thread(target=worker, daemon=True).start()

    def submit():
        input_name = name_var.get().strip()
        rrn = rrn_var.get().strip()
        sheet_name = sheet_var.get().strip()
        if not sheet_name:
            messagebox.showwarning("입력 필요", "시트를 선택하세요.")
            return
        if sheet_name not in all_sheet_names:
            messagebox.showwarning("입력 필요", "목록에서 시트를 선택하세요.")
            return
        if not input_name or not rrn:
            messagebox.showwarning("입력 필요", "이름과 주민번호를 모두 입력하세요.")
            return
        try:
            write_name_rrn(spreadsheet_id, sheet_name, input_name, rrn)
            messagebox.showinfo("완료", f"{sheet_name} 시트에 입력했습니다.")
            name_entry.focus_set()
        except Exception as e:
            messagebox.showerror("오류", f"Google Sheets API 호출 실패: {e}")

    refresh_btn = ttk.Button(frame, text="시트 새로고침", command=refresh_sheets)
    refresh_btn.grid(row=3, column=0, sticky="w", pady=8)

    submit_btn = ttk.Button(frame, text="입력", command=submit)
    submit_btn.grid(row=3, column=1, sticky="e", pady=8)

    frame.columnconfigure(1, weight=1)
    refresh_sheets()
    name_entry.focus_set()
    root.mainloop()
