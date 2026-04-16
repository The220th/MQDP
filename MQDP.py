# -*- coding: utf-8 -*-

import os
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


def resource_path(*parts: str) -> str:
    base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return str(base_dir.joinpath(*parts))


class MainWidget:
    def __init__(self, root: tk.Tk):
        self.root = root

        self.all_standards = ["standardK"]
        self.outs = ["GIFT media"]
        self.cur_standard = None

        self.standards_var = tk.StringVar(value=self.all_standards[0])
        self.path_var = tk.StringVar()
        self.out_folder_var = tk.StringVar()
        self.out_var = tk.StringVar(value=self.outs[0])

        self._build_ui()

    def _build_ui(self) -> None:
        self.root.title("MQDP")
        self.root.resizable(False, False)
        self.root.columnconfigure(0, weight=1)

        main_frame = ttk.Frame(self.root, padding=14)
        main_frame.grid(row=0, column=0, sticky="nsew")
        main_frame.columnconfigure(0, weight=1)

        ttk.Label(main_frame, text="Standards:").grid(row=0, column=0, sticky="w")

        standards_combo = ttk.Combobox(
            main_frame,
            state="readonly",
            textvariable=self.standards_var,
            values=self.all_standards,
        )
        standards_combo.grid(row=1, column=0, sticky="ew", pady=(4, 12))
        standards_combo.bind("<<ComboboxSelected>>", self._standard_combo_activated)

        ttk.Label(main_frame, text="Path to docx file:").grid(row=2, column=0, sticky="w")

        ttk.Entry(main_frame, textvariable=self.path_var).grid(
            row=3, column=0, sticky="ew", pady=(4, 4)
        )

        ttk.Button(
            main_frame,
            text="Choose docx file",
            command=self._path_interact_button_handler,
        ).grid(row=4, column=0, sticky="ew", pady=(0, 12))

        ttk.Label(main_frame, text="Out:").grid(row=5, column=0, sticky="w")

        ttk.Entry(main_frame, textvariable=self.out_folder_var).grid(
            row=6, column=0, sticky="ew", pady=(4, 4)
        )

        ttk.Button(
            main_frame,
            text="Choose out folder",
            command=self._out_folder_interact_button_handler,
        ).grid(row=7, column=0, sticky="ew", pady=(0, 4))

        ttk.Combobox(
            main_frame,
            state="readonly",
            textvariable=self.out_var,
            values=self.outs,
        ).grid(row=8, column=0, sticky="ew", pady=(0, 12))

        ttk.Button(main_frame, text="Start", command=self._start_button_handler).grid(
            row=9, column=0, sticky="ew"
        )

    def _standard_combo_activated(self, _event=None) -> None:
        self.cur_standard = self.standards_var.get()

    def _path_interact_button_handler(self) -> None:
        curdir = os.getcwd()
        filepath = filedialog.askopenfilename(
            parent=self.root,
            title="Select docx file",
            initialdir=curdir,
            filetypes=[("Docx document", "*.docx")],
        )
        if not filepath:
            return

        self.path_var.set(filepath)

        if not self.out_folder_var.get():
            out_path = filepath[: filepath.rfind(".docx")]
            self.out_folder_var.set(out_path)

    def _out_folder_interact_button_handler(self) -> None:
        curdir = os.getcwd()
        filepath = filedialog.askdirectory(
            parent=self.root,
            title="Select out folder",
            initialdir=curdir,
            mustexist=False,
        )
        if filepath:
            self.out_folder_var.set(filepath)

    def _start_button_handler(self) -> None:
        self.cur_standard = self.standards_var.get()
        if self.cur_standard is None:
            self._show_message("Please, select standard", 4)
            return

        path = self.path_var.get().strip()
        if not path:
            self._show_message("Please, set path to docx file", 4)
            return

        out_path = self.out_folder_var.get().strip()
        if not out_path:
            self._show_message("Please, select out folder", 4)
            return

        if not os.path.exists(out_path):
            os.makedirs(out_path)

        if os.path.isdir(out_path):
            if os.listdir(out_path):
                self._show_message("Out folder must be empty", 4)
                return
        else:
            self._show_message("Out folder does not exists", 4)
            return

        if self.cur_standard == self.all_standards[0]:
            from MQDP_standards import standardk_run

            res = standardk_run(path, out_path)
            self._show_message(res[0], res[1])
            return

        self._show_message("No that standard", 4)

    def _show_message(self, text: str, message_type: int) -> None:
        suffix = ""
        title = "MQDP"

        if message_type == 1:
            title = "Question"
        elif message_type in (0, 2):
            title = "Info"
        elif message_type == 3:
            title = "Warning"
        elif message_type == 4:
            title = "Error"
            suffix = "\nCheck README.md: https://github.com/The220th/MQDP/README.md"

        message = text + suffix

        if message_type == 3:
            messagebox.showwarning(title, message, parent=self.root)
        elif message_type == 4:
            messagebox.showerror(title, message, parent=self.root)
        else:
            messagebox.showinfo(title, message, parent=self.root)


def main() -> int:
    os.environ["MQPD_DEBUG_ON"] = "1"

    root = tk.Tk()

    icon_path = resource_path("imgsrc", "icon.png")
    if os.path.exists(icon_path):
        try:
            icon_image = tk.PhotoImage(file=icon_path)
            root.iconphoto(True, icon_image)
            root._icon_image = icon_image
        except tk.TclError:
            pass

    MainWidget(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
