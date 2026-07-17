# lorawan-payload-inspector

A small desktop tool for turning LoRaWAN / TTN-style JSON uplink logs into clean CSV tables and quick-look graphs — no spreadsheet gymnastics required.

Point it at a `.txt`/`.json`/`.jsonl` log where each line looks like a Things Network uplink message (`uplink_message.decoded_payload`, `received_at`, etc.), and it will:

- Parse every decoded payload field into columns
- Let you filter by time range and thin the data down to a manageable number of points
- Export exactly the columns you want as a semicolon-delimited CSV (LibreOffice/Excel-friendly)
- Plot any combination of columns as a graph — with per-line colors, auto-detected units, and separate y-axes for mismatched scales
- Show min / average / max for each plotted column, with one-click copy of just the value and unit

Built with Python's standard `tkinter` — no web server, no account, runs fully offline.

## Features

**Loading**
- Reads `.txt` / `.json` / `.jsonl` log files, one JSON object per line
- Auto-detects every payload variable across the file and lists it for selection

**Filtering**
- From/To time range, with a calendar + time picker (`YYYY/MM/DD HH:MM:SS`)
- "Max data points" thinning — evenly downsamples a huge log to a target point count instead of just truncating it

**Table export**
- Pick which columns to include
- Saves as CSV (`;`-delimited, splits into columns cleanly in LibreOffice/Excel)

**Graph export**
- Select any number of columns to plot
- Per-column line color picker and editable unit label
- **Automatic unit guessing** for common sensor fields (temperature, humidity, battery voltage vs. battery %, pressure, CO2, VOC, PM2.5/PM10, rainfall, wind speed, lux, RSSI/SNR, lat/long, pH, and more), spelling-variant tolerant (`temp`, `Temperature`, `airTemp`, `temp_c`, …) — always editable
- Option to give each column its own y-axis, or share one axis
- Option to export each selected column as its **own separate graph file** instead of one combined chart
- Choice of output format (PNG / JPG / SVG / PDF) and resolution (DPI)
- Preview opens in its own window, at the exact size/DPI you're about to export — one window per graph in "separate files" mode
- Min / average / max readout per column, with a **Copy** button on each stat that copies just `value unit` (e.g. `18.900 °C`) to the clipboard

## Requirements

```
pip install matplotlib tkcalendar
```

`tkinter` ships with most Python installs (on Linux you may need `sudo apt install python3-tk`). The calendar date picker is optional — the app still runs without `tkcalendar`, you just lose the popup picker and type dates manually.

## Usage

```
python raw_data_converter.py
```

1. **Load File** → pick your log.
2. Set a time range and/or max point count, then **Apply Filters**.
3. On the **Export Table** tab: check the columns you want, **Export to CSV**.
4. On the **Export Graph** tab: check columns, set colors/units/title, **Preview Graph**, then **Save Graph(s)**.

## Building a portable .exe (Windows)

```
pip install matplotlib tkcalendar pyinstaller
pyinstaller --onefile --windowed --name "LoRaWANPayloadInspector" raw_data_converter.py
```

The finished executable is at `dist/LoRaWANPayloadInspector.exe` — copy it anywhere, no Python install needed on the target machine. First launch is a bit slow (unpacking a `--onefile` bundle is normal). If Windows SmartScreen flags it, that's expected for an unsigned personal build.

## Notes

- Rows without a parsable `received_at` timestamp are kept for CSV export but excluded from graphs and time filtering.
- CSV output uses `;` as the delimiter (fixes LibreOffice's default comma-splitting quirk on some locales).

## License

MIT — do whatever you'd like with it.
