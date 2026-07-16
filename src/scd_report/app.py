from __future__ import annotations

from datetime import datetime
from pathlib import Path
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD

from .aggregator import build_report
from .windows import create_desktop_shortcut, resource_path


class SCDApp(TkinterDnD.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("SCD Business Report")
        self.geometry("760x330")
        self.minsize(680, 300)
        self._vars = [tk.StringVar() for _ in range(3)]
        icon = resource_path("scd.ico")
        if icon.exists():
            self.iconbitmap(default=str(icon))
        self._build()
        threading.Thread(target=self._create_shortcut, daemon=True).start()

    def _build(self) -> None:
        frame = ttk.Frame(self, padding=24)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="SCD Business Report", font=("맑은 고딕", 18, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 20))
        labels = ["팀원업무 폴더", "전체용 엑셀", "결과 저장위치"]
        commands = [self._choose_folder, self._choose_template, self._choose_output]
        for row, (label, command, var) in enumerate(zip(labels, commands, self._vars), 1):
            ttk.Label(frame, text=label, width=16).grid(row=row, column=0, sticky="w", pady=7)
            entry = ttk.Entry(frame, textvariable=var)
            entry.grid(row=row, column=1, sticky="ew", padx=8)
            entry.drop_target_register(DND_FILES)
            entry.dnd_bind("<<Drop>>", lambda event, index=row - 1: self._drop(event, index))
            ttk.Button(frame, text="선택", command=command, width=10).grid(row=row, column=2)
        self.run_button = ttk.Button(frame, text="자동취합 실행", command=self._run)
        self.run_button.grid(row=4, column=0, columnspan=3, pady=(25, 10), ipadx=35, ipady=8)
        self.status = ttk.Label(frame, text="파일과 폴더를 선택해 주세요.")
        self.status.grid(row=5, column=0, columnspan=3)
        frame.columnconfigure(1, weight=1)

    def _create_shortcut(self) -> None:
        try:
            create_desktop_shortcut()
        except Exception:
            # Shortcut creation must never prevent the report program from opening.
            pass

    def _drop(self, event, index: int) -> None:
        paths = [Path(value) for value in self.tk.splitlist(event.data)]
        if not paths:
            return
        path = paths[0]
        if index == 0:
            self._vars[0].set(str(path if path.is_dir() else path.parent))
        elif index == 1:
            if path.suffix.lower() != ".xlsx":
                messagebox.showwarning("파일 확인", "전체용 엑셀 입력란에는 .xlsx 파일을 놓아 주세요.")
                return
            self._set_template(path)
        else:
            if path.is_dir():
                path = path / f"SCD_전체업무보고_{datetime.now():%Y%m%d}.xlsx"
            elif path.suffix.lower() != ".xlsx":
                path = path.with_suffix(".xlsx")
            self._vars[2].set(str(path))

    def _set_template(self, path: Path) -> None:
        self._vars[1].set(str(path))
        if not self._vars[2].get():
            name = f"SCD_전체업무보고_{datetime.now():%Y%m%d}.xlsx"
            self._vars[2].set(str(path.with_name(name)))

    def _choose_folder(self) -> None:
        if value := filedialog.askdirectory(title="팀원업무 폴더 선택"):
            self._vars[0].set(value)

    def _choose_template(self) -> None:
        if value := filedialog.askopenfilename(title="전체용 엑셀 선택", filetypes=[("Excel", "*.xlsx")]):
            self._set_template(Path(value))

    def _choose_output(self) -> None:
        if value := filedialog.asksaveasfilename(title="결과 저장위치", defaultextension=".xlsx",
                                                 filetypes=[("Excel", "*.xlsx")]):
            self._vars[2].set(value)

    def _run(self) -> None:
        values = [v.get().strip() for v in self._vars]
        if not all(values):
            messagebox.showwarning("입력 확인", "팀원업무 폴더, 전체용 엑셀, 결과 저장위치를 모두 선택하세요.")
            return
        folder, template, output = (Path(value) for value in values)
        if not folder.is_dir() or not template.is_file():
            messagebox.showwarning("입력 확인", "선택한 팀원업무 폴더와 전체용 엑셀을 확인하세요.")
            return
        self.run_button.configure(state="disabled")
        self.status.configure(text="엑셀을 취합하고 있습니다…")
        threading.Thread(target=self._worker, args=(folder, template, output), daemon=True).start()

    def _worker(self, folder: Path, template: Path, output: Path) -> None:
        try:
            count, warnings = build_report(folder, template, output)
            self.after(0, self._done, count, warnings, output)
        except Exception as exc:
            self.after(0, self._failed, str(exc))

    def _done(self, count: int, warnings: list[str], output: Path) -> None:
        self.run_button.configure(state="normal")
        self.status.configure(text=f"완료: {count}개 업무 기록 취합")
        warning = "\n\n주의:\n" + "\n".join(warnings) if warnings else ""
        messagebox.showinfo("취합 완료", f"결과 파일을 저장했습니다.\n{output}{warning}")

    def _failed(self, message: str) -> None:
        self.run_button.configure(state="normal")
        self.status.configure(text="취합 실패")
        messagebox.showerror("오류", message)


def main() -> None:
    SCDApp().mainloop()
