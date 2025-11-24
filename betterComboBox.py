import tkinter as tk
from tkinter import ttk

class BetterComboBox(ttk.Combobox):
    def __init__(self, master=None, completevalues=None, listbox_bg="#3a3a4a", listbox_fg="white", **kwargs):
        super().__init__(master, **kwargs)
        self._completion_list = completevalues or []
        self['values'] = self._completion_list
        self._typed = ""

        self.listbox_bg = listbox_bg
        self.listbox_fg = listbox_fg

        self.bind('<KeyRelease>', self._on_keyrelease)
        self.bind('<Return>', self._show_dropdown)
        self.bind('<Escape>', self._hide_dropdown)
        self.bind('<Button-1>', self._update_listbox_style)

        # Styl comboboxa
        style = ttk.Style()
        self.custom_style_name = "Custom.TCombobox"
        self.configure(style=self.custom_style_name)

        style.configure(self.custom_style_name,
                        foreground="white",
                        background="#3a3a4a",
                        fieldbackground="#3a3a4a",
                        insertcolor="white",
                        relief="flat",
                        padding=6,
                        font=("Segoe UI", 10))

        style.map(self.custom_style_name,
                  arrowcolor=[("active", "black"), ("pressed", "black")],
                  bordercolor=[("focus", "#66ccff"), ("!focus", "#666666")])

    def _on_keyrelease(self, event):
        if event.keysym in ("Up", "Down", "Return", "Escape"):
            return
        self._typed = self.get().lower()
        filtered = [item for item in self._completion_list if self._typed in item.lower()]
        self['values'] = filtered if filtered else self._completion_list

    def set_completion_list(self, completion_list):
        self._completion_list = completion_list
        self['values'] = completion_list

    def _show_dropdown(self, event=None):
        self.event_generate('<Down>')

    def _hide_dropdown(self, event=None):
        self.selection_clear()
        self.icursor(tk.END)
        self['values'] = self._completion_list
        self.event_generate('<Escape>')

    def _update_listbox_style(self, event=None):
        """Ustawienie kolorów listy rozwijanej"""
        self.after(1, self.config_popdown)

    def config_popdown(self):
        """Zmień kolor tła i tekstu listy rozwijanej Comboboxa"""
        try:
            listbox = f"[ttk::combobox::PopdownWindow {str(self)}].f.l"
            self.tk.eval(f'{listbox} configure -background {self.listbox_bg} -foreground {self.listbox_fg}')
        except tk.TclError:
            pass  # jeśli lista się jeszcze nie otworzyła
