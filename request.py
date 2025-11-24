from tkinter import ttk, messagebox
from tkinter.ttk import Frame, Label, Button, Entry
import tkinter as tk

from opinion import Opinion
from pricing import Pricing


class Request:
    def __init__(self, window, main_frame):
        self.window = window
        self.main_frame = main_frame
        self.document_info = {}
        self.devices = None

        self.frame_request = tk.Frame(self.window, bg="#1e1e2f")
        self.frame_request.grid(row=0, column=0, sticky='nsew')
        self.frame_request.grid_rowconfigure(0, weight=1)
        self.frame_request.grid_rowconfigure(1, weight=0)
        self.frame_request.grid_columnconfigure(0, weight=1)

        # Kontener na canvas + scrollbar
        self.canvas_container = tk.Frame(self.frame_request, bg="#1e1e2f")
        self.canvas_container.grid(row=0, column=0, sticky="nsew")
        self.canvas_container.grid_rowconfigure(0, weight=1)
        self.canvas_container.grid_columnconfigure(0, weight=1)

        # Canvas
        self.devices_canvas = tk.Canvas(self.canvas_container, bg="#1e1e2f", highlightthickness=0)
        self.devices_canvas.grid(row=0, column=0, sticky="nsew")

        # Scrollbar pionowy
        self.scrollbar = ttk.Scrollbar(self.canvas_container, orient="vertical", command=self.devices_canvas.yview)
        self.scrollbar.grid(row=0, column=1, sticky="ns")

        self.devices_canvas.configure(yscrollcommand=self.scrollbar.set)

        # Wnętrze Canvas (na urządzenia)
        self.devices_container = tk.Frame(self.devices_canvas, bg="#1e1e2f")
        self.container_window = self.devices_canvas.create_window((0, 0), window=self.devices_container, anchor="nw")

        # Scrollregion i dopasowanie szerokości
        def update_scrollregion(event):
            self.devices_canvas.configure(scrollregion=self.devices_canvas.bbox("all"))

        def resize_canvas(event):
            self.devices_canvas.itemconfig(self.container_window, width=event.width)

        self.devices_container.bind("<Configure>", update_scrollregion)
        self.devices_canvas.bind("<Configure>", resize_canvas)

        # Obsługa scrolla myszką
        def _on_mousewheel(event):
            self.devices_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        self.devices_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        self.devices_canvas.bind_all("<Button-4>", lambda e: self.devices_canvas.yview_scroll(-1, "units"))
        self.devices_canvas.bind_all("<Button-5>", lambda e: self.devices_canvas.yview_scroll(1, "units"))

        self.bind_mousewheel(self.devices_container, self.devices_canvas)

        # Przyciski
        self.button_frame = tk.Frame(self.frame_request, bg="#1e1e2f")
        self.button_frame.grid(row=1, column=0, pady=20)

        self.button_add_devices = Button(self.button_frame, text="Wstecz", command=lambda: self.main_frame.tkraise())
        self.button_add_devices.pack(side="left", padx=10)

        self.button_submit = Button(self.button_frame, text="Generuj opinię", command=self.createOpinion)
        self.button_submit.pack(side="left", padx=10)

    def showRequests(self, devices):
        for widget in self.devices_container.winfo_children():
            widget.destroy()

        for idx, device in enumerate(devices):
            device.prepare_frame(self.devices_container, idx, self.document_info, par=self)

        self.devices = devices

    def createOpinion(self):
        for idx, device in enumerate(self.devices):
            if not device.field_id.get().strip():
                messagebox.showerror("Błąd danych", f"Urządzenie {idx + 1} nie ma wpisanego ID.")
                return False

            if not device.validateRequiredFields():
                return False
        for device in self.devices:
            device.updateDeviceInfo()
        try:
            self.opinion = Opinion(self.document_info, self.devices)
            messagebox.showinfo     (
                "Sukces",
                f"Opinia {self.document_info['Numer opinii']} została stworzona."
                )
        except Exception as e:
            messagebox.showerror(
                "Błąd",
                f"Nie udało się stworzyć opinii.\nSzczegóły: {e}"
            )
            self.opinion = None


        self.createPricing()

    def createPricing(self):
        for device in self.devices:
            if device.getDecision():
                try:
                    Pricing(self.document_info,device,device.pricing_file_name)
                except Exception as e:
                    messagebox.showerror(
                    "Błąd",
                    f"Nie udało się stworzyć wyceny {device.info['Numer wyceny']}.\nSzczegóły: {e}"
                )


    def bind_mousewheel(self, widget, target_canvas):
        widget.bind("<Enter>", lambda e: target_canvas.bind_all("<MouseWheel>", lambda ev: target_canvas.yview_scroll(
            int(-1 * (ev.delta / 120)), "units")))
        widget.bind("<Leave>", lambda e: target_canvas.unbind_all("<MouseWheel>"))