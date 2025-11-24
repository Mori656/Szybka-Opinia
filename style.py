from tkinter import ttk

def apply_custom_style():
    style = ttk.Style()
    style.theme_use("clam")

    style.configure("TButton",
                    background="#4a90e2",
                    foreground="white",
                    font=("Segoe UI", 10, "bold"),
                    borderwidth=0,
                    padding=10)
    style.map("TButton", background=[("active", "#357ABD")])

    style.configure("TLabel",
                    background="#1e1e2f",
                    foreground="white",
                    font=("Segoe UI", 10))

    style.configure("TEntry",
                    font=("Segoe UI", 10),
                    padding=6,
                    foreground="white",
                    fieldbackground="#3a3a4a",
                    background="#3a3a4a",
                    insertcolor="white",
                    borderwidth=1,
                    relief="flat")
    style.map("TEntry",
              foreground=[("disabled", "#666666")],
              fieldbackground=[("disabled", "#292929")],
              background=[("disabled", "#292929")])

    style.configure("Vertical.TScrollbar",
                    gripcount=0,
                    background="#1e1e2f",
                    darkcolor="#4a90e2",
                    lightcolor="#357ABD",
                    troughcolor="#1e1e2f",
                    bordercolor="#1e1e2f",
                    arrowcolor="white")
    style.map("Vertical.TScrollbar",
              background=[("active", "#357ABD"), ("pressed", "#2c5aa0")],
              arrowcolor=[("active", "white"), ("pressed", "white")])

    style.configure("TRadiobutton",
                    background="#2e2e3e",
                    foreground="white",
                    font=("Segoe UI", 10),
                    focuscolor="",
                    indicatorcolor="#4a90e2")
    style.map("TRadiobutton",
              background=[("active", "#2a3b55")],
              foreground=[("disabled", "#666666")])

    style.configure("TCheckbutton",
                    background="#2e2e3e",
                    foreground="white",
                    font=("Segoe UI", 10),
                    focuscolor="")
    style.map("TCheckbutton",
              background=[("active", "#2a3b55")],
              foreground=[("disabled", "#666666")])

    style.configure("TCombobox",
                    foreground="white",
                    background="#3a3a4a",
                    fieldbackground="#3a3a4a",
                    arrowcolor="white",
                    borderwidth=1,
                    highlightcolor = "yellow",
                    relief="flat",
                    padding=6,
                    font=("Segoe UI", 10))
    style.map("TCombobox",
              arrowcolor=[("active", "black"), ("pressed", "black")])
    style.configure("TCombobox.Listbox",
                    font=("Segoe UI", 10),
                    background="#2e2e3e",
                    foreground="white",
                    selectbackground="#4a90e2",
                    selectforeground="white")

    style.configure("Destroy.TButton",
                    background="#2e2e3e",
                    foreground="#ff5555",
                    font=("Segoe UI", 10, "bold"),
                    borderwidth=0,
                    padding=0)

    style.map("Destroy.TButton",
              background=[("active", "#661111")],
              foreground=[("active", "#ffffff")])

    style.configure("Destroy2.TButton",
                    background="#1e1e2f",
                    foreground="#ff5555",
                    font=("Segoe UI", 10, "bold"),
                    borderwidth=0,
                    padding=0)

    style.map("Destroy2.TButton",
              background=[("active", "#661111")],
              foreground=[("active", "#ffffff")])
