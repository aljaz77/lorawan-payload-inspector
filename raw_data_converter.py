import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
import json
import csv
import os
from datetime import datetime

try:
    import matplotlib
    matplotlib.use("TkAgg")
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    HAS_MPL = True
except ImportError:
    HAS_MPL = False

try:
    from tkcalendar import Calendar
    HAS_TKCALENDAR = True
except ImportError:
    HAS_TKCALENDAR = False


# Keyword -> unit lookup for automatic unit assignment. Checked in order (most
# specific/ambiguity-resolving entries first) against the variable name with
# spaces/underscores/dashes stripped and lowercased.
UNIT_RULES = [
    (["batteryvoltage", "battvolt", "vbatt", "vbat", "battv"], "V"),
    (["batterylevel", "batterypercent", "battpct", "battery", "batt"], "%"),
    (["temperature", "airtemp", "ambienttemp", "objecttemp", "tempc", "temp"], "°C"),
    (["relativehumidity", "humidity", "hum", "rhum"], "%"),
    (["pressure", "barometric", "baro", "press"], "hPa"),
    (["co2", "carbondioxide"], "ppm"),
    (["tvoc", "voc"], "ppb"),
    (["pm2.5", "pm25", "pm1.0", "pm10", "particulate", "pm"], "µg/m³"),
    (["rainfall", "precipitation", "precip", "rain"], "mm"),
    (["windspeed", "wind"], "m/s"),
    (["illuminance", "lux", "lightlevel", "light"], "lx"),
    (["distance", "range"], "m"),
    (["altitude", "elevation"], "m"),
    (["latitude"], "°"),
    (["longitude", "lng"], "°"),
    (["rssi"], "dBm"),
    (["snr"], "dB"),
    (["amperage", "amps", "current"], "A"),
    (["voltage", "volt"], "V"),
    (["power", "watt"], "W"),
    (["energy", "kwh"], "kWh"),
    (["frequency", "freq", "hertz"], "Hz"),
    (["angle", "tilt", "orientation"], "°"),
    (["velocity", "speed"], "m/s"),
    (["weight", "mass"], "kg"),
    (["soilmoisture", "moisture"], "%"),
    (["noise", "sound", "decibel"], "dB"),
    (["conductivity"], "µS/cm"),
    (["phlevel", "ph"], "pH"),
    (["counter", "pulses", "count"], "count"),
    (["uptime", "duration"], "s"),
]

CHART_COLOR_PALETTE = ["#1f77b4", "#d62728", "#2ca02c", "#9467bd", "#ff7f0e",
                        "#17becf", "#e377c2", "#8c564b", "#bcbd22", "#7f7f7f"]


def auto_unit(var_name):
    normalized = var_name.lower().replace("_", "").replace("-", "").replace(" ", "")
    for keywords, unit in UNIT_RULES:
        for kw in keywords:
            if kw in normalized:
                return unit
    return ""


class SensorDataConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Network Payload to Spreadsheet Converter")
        self.root.geometry("680x820")

        # ---- data state ----
        self.data = []            # every parsed row (dict), includes "_dt" datetime or None
        self.filtered_data = []   # after time-range filter + downsampling
        self.variables = []       # sorted list of payload variable names
        self.checkbox_vars = {}   # var name -> BooleanVar (table export column selection)
        self.graph_vars = {}      # var name -> BooleanVar (graph column selection)
        self.graph_units = {}     # var name -> StringVar (unit label for that column)
        self.graph_colors = {}    # var name -> StringVar (hex color for that column's line)

        # ================= Top: load file =================
        top = tk.Frame(root, padx=15, pady=10)
        top.pack(fill=tk.X)

        tk.Label(top, text="Step 1: Load log file (.txt, .json, .jsonl)", font=("Arial", 10, "bold")).pack(anchor="w")
        row = tk.Frame(top)
        row.pack(fill=tk.X, pady=5)
        self.load_btn = tk.Button(row, text="Load File", command=self.load_file, width=18)
        self.load_btn.pack(side=tk.LEFT)
        self.file_label = tk.Label(row, text="No file loaded", fg="gray")
        self.file_label.pack(side=tk.LEFT, padx=10)

        # ================= Filter bar (time range + downsample) =================
        filt = tk.LabelFrame(root, text="Step 2: Filter & reduce data", padx=15, pady=10)
        filt.pack(fill=tk.X, padx=15, pady=(0, 10))

        tk.Label(filt, text="From:").grid(row=0, column=0, sticky="w")
        self.from_entry = tk.Entry(filt, width=20)
        self.from_entry.grid(row=0, column=1, padx=5, pady=3)
        tk.Button(filt, text="📅", width=3, command=lambda: self._open_date_picker(self.from_entry)).grid(
            row=0, column=2, padx=(0, 15))

        tk.Label(filt, text="To:").grid(row=0, column=3, sticky="w")
        self.to_entry = tk.Entry(filt, width=20)
        self.to_entry.grid(row=0, column=4, padx=5, pady=3)
        tk.Button(filt, text="📅", width=3, command=lambda: self._open_date_picker(self.to_entry)).grid(
            row=0, column=5)

        tk.Label(filt, text="(format: YYYY/MM/DD HH:MM:SS)", fg="gray", font=("Arial", 8)).grid(
            row=1, column=0, columnspan=6, sticky="w")

        tk.Label(filt, text="Max data points (0 = no limit, thins evenly):").grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))
        self.max_points_entry = tk.Entry(filt, width=10)
        self.max_points_entry.insert(0, "0")
        self.max_points_entry.grid(row=2, column=3, sticky="w", pady=(8, 0))

        self.apply_btn = tk.Button(filt, text="Apply Filters", command=self.apply_filters, state=tk.DISABLED)
        self.apply_btn.grid(row=2, column=4, columnspan=2, sticky="e", pady=(8, 0))

        self.filter_status = tk.Label(filt, text="No data loaded yet.", fg="gray")
        self.filter_status.grid(row=3, column=0, columnspan=6, sticky="w", pady=(6, 0))

        # ================= Tabs: Table export / Graph export =================
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 10))

        self.table_tab = tk.Frame(self.notebook)
        self.graph_tab = tk.Frame(self.notebook)
        self.notebook.add(self.table_tab, text="Export Table (CSV)")
        self.notebook.add(self.graph_tab, text="Export Graph (Image)")

        self._build_table_tab()
        self._build_graph_tab()

    # ======================================================================
    # Loading
    # ======================================================================
    def load_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Text/JSON Files", "*.txt *.json *.jsonl"), ("All Files", "*.*")])
        if not filepath:
            return

        self.data = []
        var_set = set()

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    try:
                        entry = json.loads(line)

                        raw_time = entry.get("received_at", "")
                        date_str = ""
                        time_str = ""
                        dt_obj = None
                        if raw_time:
                            try:
                                clean_time = raw_time.split('.')[0].replace('Z', '')
                                dt_obj = datetime.fromisoformat(clean_time)
                                date_str = dt_obj.strftime("%Y-%m-%d")
                                time_str = dt_obj.strftime("%H:%M:%S")
                            except ValueError:
                                date_str = raw_time

                        payload = entry.get("uplink_message", {}).get("decoded_payload", {})
                        if not payload:
                            continue

                        row = {"Date": date_str, "Time": time_str, "_dt": dt_obj}
                        for k, v in payload.items():
                            row[k] = v
                            var_set.add(k)

                        self.data.append(row)
                    except json.JSONDecodeError:
                        continue

            self.data.sort(key=lambda r: (r["_dt"] is None, r["_dt"]))
            self.variables = sorted(list(var_set))

            self._populate_column_lists()
            self._populate_time_bounds()

            self.filtered_data = list(self.data)
            filename = filepath.split('/')[-1]
            self.file_label.config(text=f"Loaded: {filename}  ({len(self.data)} rows)", fg="green")
            self.apply_btn.config(state=tk.NORMAL)
            self.filter_status.config(text=f"{len(self.data)} of {len(self.data)} rows selected.", fg="black")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to read the file:\n{e}")

    def _populate_time_bounds(self):
        dts = [r["_dt"] for r in self.data if r["_dt"] is not None]
        self.from_entry.delete(0, tk.END)
        self.to_entry.delete(0, tk.END)
        if dts:
            self.from_entry.insert(0, min(dts).strftime("%Y/%m/%d %H:%M:%S"))
            self.to_entry.insert(0, max(dts).strftime("%Y/%m/%d %H:%M:%S"))

    def _open_date_picker(self, entry_widget):
        if not HAS_TKCALENDAR:
            messagebox.showinfo(
                "Calendar picker unavailable",
                "The calendar picker needs an extra package.\n\nInstall it with:\n\npip install tkcalendar"
            )
            return

        # seed the picker from whatever is currently in the entry, if valid
        current = entry_widget.get().strip()
        init_h, init_m, init_s = 0, 0, 0
        cal_kwargs = {"date_pattern": "yyyy/mm/dd"}
        if current:
            try:
                dt = datetime.strptime(current, "%Y/%m/%d %H:%M:%S")
                cal_kwargs.update(year=dt.year, month=dt.month, day=dt.day)
                init_h, init_m, init_s = dt.hour, dt.minute, dt.second
            except ValueError:
                pass

        popup = tk.Toplevel(self.root)
        popup.title("Pick date & time")
        popup.grab_set()

        cal = Calendar(popup, selectmode="day", **cal_kwargs)
        cal.pack(padx=10, pady=10)

        time_row = tk.Frame(popup)
        time_row.pack(pady=(0, 10))
        tk.Label(time_row, text="Time:").pack(side=tk.LEFT)
        h_var = tk.StringVar(value=f"{init_h:02d}")
        m_var = tk.StringVar(value=f"{init_m:02d}")
        s_var = tk.StringVar(value=f"{init_s:02d}")
        tk.Spinbox(time_row, from_=0, to=23, width=3, format="%02.0f", textvariable=h_var).pack(side=tk.LEFT)
        tk.Label(time_row, text=":").pack(side=tk.LEFT)
        tk.Spinbox(time_row, from_=0, to=59, width=3, format="%02.0f", textvariable=m_var).pack(side=tk.LEFT)
        tk.Label(time_row, text=":").pack(side=tk.LEFT)
        tk.Spinbox(time_row, from_=0, to=59, width=3, format="%02.0f", textvariable=s_var).pack(side=tk.LEFT)

        def confirm():
            try:
                h, m, s = int(h_var.get()), int(m_var.get()), int(s_var.get())
            except ValueError:
                messagebox.showerror("Error", "Time must be numeric.")
                return
            final = f"{cal.get_date()} {h:02d}:{m:02d}:{s:02d}"
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, final)
            popup.destroy()

        tk.Button(popup, text="Set", command=confirm, bg="#4CAF50", fg="white", width=10).pack(pady=(0, 10))

    def _populate_column_lists(self):
        # table tab checkboxes
        for widget in self.table_check_frame.winfo_children():
            widget.destroy()
        self.checkbox_vars.clear()
        for var in self.variables:
            v = tk.BooleanVar(value=True)
            self.checkbox_vars[var] = v
            tk.Checkbutton(self.table_check_frame, text=var, variable=v, font=("Arial", 10)).pack(anchor="w", padx=5, pady=1)

        # graph tab checkboxes + unit entries + color pickers
        for widget in self.graph_check_frame.winfo_children():
            widget.destroy()
        self.graph_vars.clear()
        self.graph_units.clear()
        self.graph_colors.clear()
        for i, var in enumerate(self.variables):
            v = tk.BooleanVar(value=False)
            u = tk.StringVar(value=auto_unit(var))
            c = tk.StringVar(value=CHART_COLOR_PALETTE[i % len(CHART_COLOR_PALETTE)])
            self.graph_vars[var] = v
            self.graph_units[var] = u
            self.graph_colors[var] = c

            r = tk.Frame(self.graph_check_frame)
            r.pack(fill=tk.X, padx=5, pady=1)
            tk.Checkbutton(r, text=var, variable=v, font=("Arial", 10), width=20, anchor="w").pack(side=tk.LEFT)
            tk.Label(r, text="unit:", fg="gray").pack(side=tk.LEFT, padx=(5, 2))
            tk.Entry(r, textvariable=u, width=8).pack(side=tk.LEFT)

            swatch = tk.Button(r, text="  ", bg=c.get(), width=2,
                                command=lambda var=var: self._choose_color(var))
            swatch.pack(side=tk.LEFT, padx=(8, 0))
            self._color_swatches = getattr(self, "_color_swatches", {})
            self._color_swatches[var] = swatch

    def _bind_mousewheel(self, canvas):
        """Let the mouse wheel scroll this canvas while the cursor is over it
        (Windows/Mac use <MouseWheel>, Linux uses Button-4/5)."""
        def _on_mousewheel(event):
            if event.num == 4:
                canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                canvas.yview_scroll(1, "units")
            else:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _bind_all(_event=None):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            canvas.bind_all("<Button-4>", _on_mousewheel)
            canvas.bind_all("<Button-5>", _on_mousewheel)

        def _unbind_all(_event=None):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        canvas.bind("<Enter>", _bind_all)
        canvas.bind("<Leave>", _unbind_all)

    def _choose_color(self, var):
        current = self.graph_colors[var].get()
        rgb, hexval = colorchooser.askcolor(color=current, title=f"Line color for {var}")
        if hexval:
            self.graph_colors[var].set(hexval)
            self._color_swatches[var].config(bg=hexval)

    # ======================================================================
    # Filtering / downsampling
    # ======================================================================
    def apply_filters(self):
        if not self.data:
            return

        from_str = self.from_entry.get().strip()
        to_str = self.to_entry.get().strip()

        try:
            from_dt = datetime.strptime(from_str, "%Y/%m/%d %H:%M:%S") if from_str else None
            to_dt = datetime.strptime(to_str, "%Y/%m/%d %H:%M:%S") if to_str else None
        except ValueError:
            messagebox.showerror("Error", "Time filters must be in format YYYY/MM/DD HH:MM:SS")
            return

        rows = []
        for r in self.data:
            dt = r["_dt"]
            if dt is None:
                continue
            if from_dt and dt < from_dt:
                continue
            if to_dt and dt > to_dt:
                continue
            rows.append(r)

        try:
            max_points = int(self.max_points_entry.get().strip() or "0")
        except ValueError:
            messagebox.showerror("Error", "Max data points must be a whole number.")
            return

        if max_points > 0 and len(rows) > max_points:
            stride = len(rows) / max_points
            thinned = []
            i = 0.0
            while int(i) < len(rows):
                thinned.append(rows[int(i)])
                i += stride
            rows = thinned

        self.filtered_data = rows
        self.filter_status.config(text=f"{len(rows)} of {len(self.data)} rows selected.", fg="black")

    # ======================================================================
    # Table tab (CSV export)
    # ======================================================================
    def _build_table_tab(self):
        tk.Label(self.table_tab, text="Select columns to export", font=("Arial", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 0))

        container = tk.Frame(self.table_tab, relief=tk.SUNKEN, borderwidth=1)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.table_check_frame = tk.Frame(canvas)

        self.table_check_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.table_check_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._bind_mousewheel(canvas)

        self.export_btn = tk.Button(self.table_tab, text="Export to CSV", command=self.export_file,
                                     width=20, bg="#4CAF50", fg="white")
        self.export_btn.pack(pady=10)

    def export_file(self):
        if not self.filtered_data:
            messagebox.showwarning("Warning", "No data to export. Load a file and apply filters first.")
            return

        selected_vars = [var for var, is_checked in self.checkbox_vars.items() if is_checked.get()]
        if not selected_vars:
            messagebox.showwarning("Warning", "Please select at least one variable.")
            return

        headers = ["Date", "Time"] + selected_vars

        filepath = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV File", "*.csv")])
        if not filepath:
            return

        try:
            with open(filepath, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=headers, extrasaction='ignore', delimiter=';')
                writer.writeheader()
                for row in self.filtered_data:
                    writer.writerow(row)
            messagebox.showinfo("Success", "Data exported!\n\nOpen this in LibreOffice. It should now split into columns automatically.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save the file:\n{e}")

    # ======================================================================
    # Graph tab
    # ======================================================================
    def _build_graph_tab(self):
        if not HAS_MPL:
            tk.Label(self.graph_tab, text="matplotlib is not installed.\nRun: pip install matplotlib",
                     fg="red", font=("Arial", 10, "bold")).pack(pady=30)
            return

        tk.Label(self.graph_tab, text="Select column(s) to plot - set a unit and line color for each",
                 font=("Arial", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 0))

        container = tk.Frame(self.graph_tab, relief=tk.SUNKEN, borderwidth=1, height=150)
        container.pack(fill=tk.X, padx=10, pady=5)
        container.pack_propagate(False)

        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.graph_check_frame = tk.Frame(canvas)

        self.graph_check_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.graph_check_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._bind_mousewheel(canvas)

        opts = tk.LabelFrame(self.graph_tab, text="Chart text & layout", padx=10, pady=8)
        opts.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(opts, text="Title:").grid(row=0, column=0, sticky="w")
        self.graph_title_entry = tk.Entry(opts, width=45)
        self.graph_title_entry.insert(0, "Sensor Data")
        self.graph_title_entry.grid(row=0, column=1, padx=5, pady=3, sticky="w")

        tk.Label(opts, text="Note (bottom of chart):").grid(row=1, column=0, sticky="w")
        self.graph_note_entry = tk.Entry(opts, width=45)
        self.graph_note_entry.grid(row=1, column=1, padx=5, pady=3, sticky="w")

        self.separate_axes_var = tk.BooleanVar(value=True)
        tk.Checkbutton(opts, text="Give each selected column its own y-axis (combined chart)",
                        variable=self.separate_axes_var).grid(row=2, column=0, columnspan=2, sticky="w", pady=(5, 0))

        self.separate_files_var = tk.BooleanVar(value=False)
        tk.Checkbutton(opts, text="Export each selected column as its own separate graph/file",
                        variable=self.separate_files_var).grid(row=3, column=0, columnspan=2, sticky="w")

        tk.Label(opts, text="Image format:").grid(row=4, column=0, sticky="w", pady=(8, 0))
        self.image_format_var = tk.StringVar(value="PNG")
        format_menu = tk.OptionMenu(opts, self.image_format_var, "PNG", "JPG", "SVG", "PDF")
        format_menu.config(width=8)
        format_menu.grid(row=4, column=1, sticky="w", pady=(8, 0))

        tk.Label(opts, text="Resolution (DPI):").grid(row=5, column=0, sticky="w")
        self.dpi_entry = tk.Entry(opts, width=10)
        self.dpi_entry.insert(0, "150")
        self.dpi_entry.grid(row=5, column=1, sticky="w")
        tk.Label(opts, text="(image sharpness - higher = larger file, ignored for SVG/PDF)",
                 fg="gray", font=("Arial", 8)).grid(row=6, column=0, columnspan=2, sticky="w")

        btns = tk.Frame(self.graph_tab)
        btns.pack(pady=8)
        tk.Button(btns, text="Preview Graph", command=self.preview_graph, width=18).pack(side=tk.LEFT, padx=5)
        self.save_graph_btn = tk.Button(btns, text="Save Graph(s)", command=self.save_graph,
                                         width=20, bg="#4CAF50", fg="white")
        self.save_graph_btn.pack(side=tk.LEFT, padx=5)

        tk.Label(self.graph_tab, text="Previews open in their own window(s), at the format/resolution set above.",
                 fg="gray", font=("Arial", 8)).pack(pady=(0, 5))

        # Min / average / max readout, with a per-value copy button
        stats_outer = tk.LabelFrame(self.graph_tab, text="Min / Average / Max", padx=10, pady=6)
        stats_outer.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        stats_container = tk.Frame(stats_outer, relief=tk.SUNKEN, borderwidth=1)
        stats_container.pack(fill=tk.BOTH, expand=True)

        stats_canvas = tk.Canvas(stats_container, highlightthickness=0)
        stats_scroll = tk.Scrollbar(stats_container, orient="vertical", command=stats_canvas.yview)
        self.stats_frame = tk.Frame(stats_canvas)

        self.stats_frame.bind("<Configure>", lambda e: stats_canvas.configure(scrollregion=stats_canvas.bbox("all")))
        stats_canvas.create_window((0, 0), window=self.stats_frame, anchor="nw")
        stats_canvas.configure(yscrollcommand=stats_scroll.set)
        stats_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        stats_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._bind_mousewheel(stats_canvas)

        self._preview_windows = []

    def _selected_graph_vars(self):
        return [v for v, checked in self.graph_vars.items() if checked.get()]

    def _numeric_series(self, var, rows=None):
        """Return (times, values) of numeric datapoints for var, using filtered_data by default."""
        rows = self.filtered_data if rows is None else rows
        times, values = [], []
        for r in rows:
            val = r.get(var)
            if isinstance(val, (int, float)) and r.get("_dt") is not None:
                times.append(r["_dt"])
                values.append(val)
        return times, values

    def _copy_to_clipboard(self, text):
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update()

    def _add_stat_row(self, parent, label, value, unit):
        """One row: 'label: value unit  at timestamp' plus a Copy button that
        copies only 'value unit' to the clipboard."""
        copy_text = f"{value:.3f}{(' ' + unit) if unit else ''}"
        row = tk.Frame(parent)
        row.pack(fill=tk.X, anchor="w", pady=1)
        tk.Label(row, text=label, font=("Consolas", 9), width=42, anchor="w").pack(side=tk.LEFT)
        tk.Button(row, text="Copy", width=6, command=lambda: self._copy_to_clipboard(copy_text)).pack(side=tk.LEFT, padx=4)

    def _update_stats(self, selected_vars):
        for widget in self.stats_frame.winfo_children():
            widget.destroy()

        for var in selected_vars:
            times, values = self._numeric_series(var)
            unit = self.graph_units[var].get().strip()

            tk.Label(self.stats_frame, text=var, font=("Consolas", 9, "bold"), anchor="w").pack(
                fill=tk.X, anchor="w", pady=(6, 0))

            if not values:
                tk.Label(self.stats_frame, text="   no numeric data", fg="gray", font=("Consolas", 9)).pack(anchor="w")
                continue

            min_idx = values.index(min(values))
            max_idx = values.index(max(values))
            vmin, vmax = values[min_idx], values[max_idx]
            vavg = sum(values) / len(values)
            min_time = times[min_idx].strftime("%Y/%m/%d %H:%M:%S")
            max_time = times[max_idx].strftime("%Y/%m/%d %H:%M:%S")
            unit_suffix = f" {unit}" if unit else ""

            self._add_stat_row(self.stats_frame, f"   min: {vmin:.3f}{unit_suffix}   at {min_time}", vmin, unit)
            self._add_stat_row(self.stats_frame, f"   avg: {vavg:.3f}{unit_suffix}", vavg, unit)
            self._add_stat_row(self.stats_frame, f"   max: {vmax:.3f}{unit_suffix}   at {max_time}", vmax, unit)

    def _build_combined_figure(self, selected_vars, dpi):
        rows = [r for r in self.filtered_data if r.get("_dt") is not None]
        if not rows:
            messagebox.showwarning("Warning", "No timestamped rows available to plot.")
            return None

        fig = Figure(figsize=(7.5, 4.5), dpi=dpi)
        ax_main = fig.add_subplot(111)

        lines, labels = [], []
        separate = self.separate_axes_var.get() and len(selected_vars) > 1

        for i, var in enumerate(selected_vars):
            xs, ys = self._numeric_series(var, rows)
            if not xs:
                continue

            unit = self.graph_units[var].get().strip()
            axis_label = f"{var} ({unit})" if unit else var
            color = self.graph_colors[var].get()

            if i == 0:
                ax = ax_main
            elif separate:
                ax = ax_main.twinx()
                if i > 1:
                    ax.spines["right"].set_position(("outward", 60 * (i - 1)))
            else:
                ax = ax_main

            line, = ax.plot(xs, ys, label=axis_label, color=color, marker="o", markersize=3, linewidth=1.2)
            lines.append(line)
            labels.append(axis_label)

            if separate or len(selected_vars) == 1:
                ax.set_ylabel(axis_label, color=color)
                ax.tick_params(axis="y", labelcolor=color)
            elif i == 0:
                ax.set_ylabel("Value")

        ax_main.set_xlabel("Time")
        title = self.graph_title_entry.get().strip()
        if title:
            ax_main.set_title(title)
        if not separate and len(selected_vars) > 1:
            ax_main.legend(lines, labels, loc="best")

        note = self.graph_note_entry.get().strip()
        if note:
            fig.text(0.5, 0.01, note, ha="center", fontsize=8, style="italic")

        fig.autofmt_xdate()
        fig.tight_layout(rect=[0, 0.03, 1, 1])
        return fig

    def _build_single_column_figure(self, var, dpi):
        xs, ys = self._numeric_series(var)
        if not xs:
            return None

        unit = self.graph_units[var].get().strip()
        axis_label = f"{var} ({unit})" if unit else var
        color = self.graph_colors[var].get()

        fig = Figure(figsize=(7.5, 4.5), dpi=dpi)
        ax = fig.add_subplot(111)
        ax.plot(xs, ys, label=axis_label, color=color, marker="o", markersize=3, linewidth=1.2)
        ax.set_xlabel("Time")
        ax.set_ylabel(axis_label, color=color)
        ax.tick_params(axis="y", labelcolor=color)

        base_title = self.graph_title_entry.get().strip()
        ax.set_title(f"{base_title} - {var}" if base_title else var)

        note = self.graph_note_entry.get().strip()
        if note:
            fig.text(0.5, 0.01, note, ha="center", fontsize=8, style="italic")

        fig.autofmt_xdate()
        fig.tight_layout(rect=[0, 0.03, 1, 1])
        return fig

    def _close_preview_windows(self):
        for win in self._preview_windows:
            try:
                win.destroy()
            except tk.TclError:
                pass
        self._preview_windows = []

    def _show_figure_window(self, fig, title):
        win = tk.Toplevel(self.root)
        win.title(title)
        canvas = FigureCanvasTkAgg(fig, master=win)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self._preview_windows.append(win)

    def preview_graph(self):
        selected_vars = self._selected_graph_vars()
        if not selected_vars:
            messagebox.showwarning("Warning", "Select at least one column to plot.")
            return
        if not self.filtered_data:
            messagebox.showwarning("Warning", "No data to plot. Load a file and apply filters first.")
            return

        try:
            dpi = int(self.dpi_entry.get().strip() or "150")
        except ValueError:
            messagebox.showerror("Error", "Resolution (DPI) must be a whole number.")
            return

        self._close_preview_windows()

        if self.separate_files_var.get():
            skipped = []
            for var in selected_vars:
                fig = self._build_single_column_figure(var, dpi)
                if fig is None:
                    skipped.append(var)
                    continue
                self._show_figure_window(fig, title=var)
            if skipped and len(skipped) == len(selected_vars):
                messagebox.showwarning("Warning", "No numeric data to plot for the selected column(s).")
        else:
            fig = self._build_combined_figure(selected_vars, dpi)
            if fig is None:
                return
            self._show_figure_window(fig, title=self.graph_title_entry.get().strip() or "Graph Preview")

        self._update_stats(selected_vars)

    def save_graph(self):
        selected_vars = self._selected_graph_vars()
        if not selected_vars:
            messagebox.showwarning("Warning", "Select at least one column to plot.")
            return

        fmt = self.image_format_var.get().lower()
        ext = f".{fmt}"
        filetype_labels = {
            "png": ("PNG Image", "*.png"),
            "jpg": ("JPEG Image", "*.jpg"),
            "svg": ("SVG Image", "*.svg"),
            "pdf": ("PDF Document", "*.pdf"),
        }

        try:
            dpi = int(self.dpi_entry.get().strip() or "150")
        except ValueError:
            messagebox.showerror("Error", "Resolution (DPI) must be a whole number.")
            return

        if self.separate_files_var.get():
            folder = filedialog.askdirectory(title="Choose folder to save the graphs into")
            if not folder:
                return

            saved, skipped = [], []
            for var in selected_vars:
                fig = self._build_single_column_figure(var, dpi)
                if fig is None:
                    skipped.append(var)
                    continue
                safe_name = "".join(c if c.isalnum() or c in "-_" else "_" for c in var)
                out_path = os.path.join(folder, f"{safe_name}{ext}")
                fig.savefig(out_path, dpi=dpi, bbox_inches="tight")
                saved.append(out_path)

            self._update_stats(selected_vars)

            msg = f"Saved {len(saved)} graph(s) to:\n{folder}"
            if skipped:
                msg += f"\n\nSkipped (no numeric data): {', '.join(skipped)}"
            messagebox.showinfo("Success", msg)
        else:
            fig = self._build_combined_figure(selected_vars, dpi)
            if fig is None:
                return
            filepath = filedialog.asksaveasfilename(defaultextension=ext, filetypes=[filetype_labels[fmt]])
            if not filepath:
                return
            try:
                fig.savefig(filepath, dpi=dpi, bbox_inches="tight")
                messagebox.showinfo("Success", "Graph saved!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save the graph:\n{e}")

            self._update_stats(selected_vars)


if __name__ == "__main__":
    root = tk.Tk()
    app = SensorDataConverter(root)
    root.mainloop()
