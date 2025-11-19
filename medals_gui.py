import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import List

import make_medals


class MedalsApp:
    def __init__(self, master: tk.Tk) -> None:
        self.master = master
        master.title("Генератор грамот")
        master.geometry("720x620")

        # Tk variables
        self.template_var = tk.StringVar(value="medals.docx")
        self.excel_var = tk.StringVar(value="listm.xlsx")
        self.sheet_var = tk.StringVar()
        self.output_var = tk.StringVar(value="medals_out.docx")
        self.outdir_var = tk.StringVar(value="out")
        self.mode_var = tk.StringVar(value="separate")
        self.placeholder_surname_var = tk.StringVar(value=make_medals.DEFAULT_SURNAME_PLACEHOLDER)
        self.placeholder_name_var = tk.StringVar(value=make_medals.DEFAULT_NAME_PLACEHOLDER)
        self.placeholder_patronymic_var = tk.StringVar(value=make_medals.DEFAULT_PATRONYMIC_PLACEHOLDER)

        self.placeholders_text = tk.Text(master, height=5)
        default_placeholder_text = "\n".join(
            [make_medals.DEFAULT_FULL_PLACEHOLDER, make_medals.DEFAULT_FULL_PLACEHOLDER_NOM]
        )
        self.placeholders_text.insert("1.0", default_placeholder_text)

        self.status_var = tk.StringVar(value="Готово до запуску")

        self._build_layout()

    # UI builders
    def _build_layout(self) -> None:
        frm_paths = ttk.LabelFrame(self.master, text="Файли")
        frm_paths.pack(fill="x", padx=10, pady=8)

        self._build_path_row(frm_paths, "Шаблон DOCX", self.template_var, self._browse_template)
        self._build_path_row(frm_paths, "Excel з ПІБ", self.excel_var, self._browse_excel)

        frm_options = ttk.LabelFrame(self.master, text="Параметри")
        frm_options.pack(fill="x", padx=10, pady=8)

        self._add_entry(frm_options, "Аркуш Excel", self.sheet_var, row=0)
        self._add_entry(frm_options, "Вихідний DOCX", self.output_var, row=1)
        self._add_entry(frm_options, "Каталог окремих файлів", self.outdir_var, row=2)

        frm_mode = ttk.Frame(frm_options)
        frm_mode.grid(column=0, row=3, columnspan=3, sticky="w", pady=4)
        ttk.Label(frm_mode, text="Режим роботи:").pack(side="left", padx=(0, 10))
        ttk.Radiobutton(frm_mode, text="Окремі файли", value="separate", variable=self.mode_var).pack(side="left")
        ttk.Radiobutton(frm_mode, text="Єдиний файл", value="single", variable=self.mode_var).pack(
            side="left", padx=10
        )

        frm_placeholders = ttk.LabelFrame(self.master, text="Плейсхолдери")
        frm_placeholders.pack(fill="both", padx=10, pady=8, expand=True)

        ttk.Label(
            frm_placeholders, text="Список повних плейсхолдерів (кожен із нового рядка)"
        ).pack(anchor="w")
        self.placeholders_text.pack(fill="both", padx=6, pady=4)

        grid = ttk.Frame(frm_placeholders)
        grid.pack(fill="x", padx=6, pady=4)
        self._add_entry(grid, "Плейсхолдер прізвища", self.placeholder_surname_var, row=0)
        self._add_entry(grid, "Плейсхолдер імені", self.placeholder_name_var, row=1)
        self._add_entry(grid, "Плейсхолдер по батькові", self.placeholder_patronymic_var, row=2)

        frm_actions = ttk.Frame(self.master)
        frm_actions.pack(fill="x", padx=10, pady=10)
        self.run_button = ttk.Button(frm_actions, text="Запустити", command=self.run_generation)
        self.run_button.pack(side="left")
        ttk.Label(frm_actions, textvariable=self.status_var).pack(side="left", padx=12)

    def _build_path_row(self, parent, label, variable, command) -> None:
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=4)
        ttk.Label(row, text=label, width=18).pack(side="left")
        entry = ttk.Entry(row, textvariable=variable)
        entry.pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(row, text="Огляд", command=command).pack(side="left")

    def _add_entry(self, parent, label, variable, row=0) -> None:
        ttk.Label(parent, text=label).grid(column=0, row=row, sticky="w", pady=2, padx=2)
        entry = ttk.Entry(parent, textvariable=variable)
        entry.grid(column=1, row=row, sticky="we", padx=2, pady=2)
        parent.columnconfigure(1, weight=1)

    # Browsers
    def _browse_template(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("DOCX", "*.docx"), ("Усі файли", "*.*")])
        if path:
            self.template_var.set(path)

    def _browse_excel(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx"), ("Усі файли", "*.*")])
        if path:
            self.excel_var.set(path)

    # Helpers
    def _collect_placeholders(self) -> List[str]:
        text = self.placeholders_text.get("1.0", "end").strip()
        return [line.strip() for line in text.splitlines() if line.strip()]

    def build_args(self) -> List[str]:
        args: List[str] = [
            "--template",
            self.template_var.get().strip(),
            "--excel",
            self.excel_var.get().strip(),
            "--output",
            self.output_var.get().strip() or "medals_out.docx",
            "--out-dir",
            self.outdir_var.get().strip() or "out",
        ]
        sheet = self.sheet_var.get().strip()
        if sheet:
            args.extend(["--sheet", sheet])
        placeholders = self._collect_placeholders()
        for ph in placeholders:
            args.extend(["--placeholder", ph])
        surname_ph = self.placeholder_surname_var.get().strip()
        if surname_ph:
            args.extend(["--placeholder-surname", surname_ph])
        name_ph = self.placeholder_name_var.get().strip()
        if name_ph:
            args.extend(["--placeholder-name", name_ph])
        patronymic_ph = self.placeholder_patronymic_var.get().strip()
        if patronymic_ph:
            args.extend(["--placeholder-patronymic", patronymic_ph])
        if self.mode_var.get() == "single":
            args.append("--single")
        else:
            args.append("--separate")
        return args

    def run_generation(self) -> None:
        args = self.build_args()
        self.run_button.config(state="disabled")
        self.status_var.set("Виконується...")

        def worker() -> None:
            try:
                code = make_medals.main(args)
            except Exception as exc:  # pragma: no cover - GUI feedback
                self.master.after(0, self._handle_result, exc)
                return
            self.master.after(0, self._handle_result, code)

        threading.Thread(target=worker, daemon=True).start()

    def _handle_result(self, result) -> None:
        self.run_button.config(state="normal")
        if isinstance(result, Exception):
            self.status_var.set("Сталася помилка")
            messagebox.showerror("Помилка", str(result))
            return
        if result == 0:
            self.status_var.set("Готово")
            messagebox.showinfo("Готово", "Генерацію завершено успішно.")
        else:
            self.status_var.set("Завершено з помилками")
            messagebox.showwarning(
                "Попередження", f"Скрипт завершився з кодом {result}. Перевірте консоль."
            )


def main() -> None:
    root = tk.Tk()
    MedalsApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
