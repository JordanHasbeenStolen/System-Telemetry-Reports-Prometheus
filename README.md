# System-Telemetry-Reports-Prometheus

Generates structured Excel reports from Prometheus metrics, providing a detailed overview of PC hardware telemetry, performance, temperatures, resource utilization, and system health for anomaly detection during browser workload testing.

The script pulls the full history of `node_exporter` (and optionally NVIDIA GPU) metrics from a running Prometheus instance over a chosen time window and writes them into a multi-sheet `.xlsx` file, one sheet per hardware category (CPU, Temperature, Memory, Disk, Network, System, Power, GPU, Other).

---

## What it does

- Queries Prometheus `query_range` API for every `node_*` metric over the last N hours.
- Buckets each metric into a themed sheet by prefix (CPU, Temperature, Memory, Disk, Network, System, Power, Other).
- Optionally collects NVIDIA GPU metrics (`nvidia_*`) into a dedicated **GPU** sheet.
- Writes a styled, human-readable Excel workbook with a **Summary** sheet showing row counts per section.
- Automatically splits oversized sheets (Excel caps at 1,048,576 rows) into `Section_2`, `Section_3`, etc.

---

## Requirements

### System services
This script is only the *export* layer. It expects a working monitoring stack:

1. **Prometheus** — running and reachable at `http://localhost:9090`.
2. **node_exporter** — running (default port `9100`), scraped by Prometheus.
3. **(Optional) nvidia_gpu_exporter** — for GPU metrics (default port `9835`).
4. **(Optional) Grafana** — not required by the script, but useful for live dashboards.

### Python
- Python 3.8+
- Packages:
  - `requests`
  - `openpyxl`
  - `pandas` (imported; safe to keep installed)

Install:
```bash
pip install requests openpyxl pandas
```

---

## Installing the monitoring stack (Ubuntu / Linux Mint)

### 1. Prometheus
```bash
sudo apt update
sudo apt install prometheus
```
Config lives at `/etc/prometheus/prometheus.yml`. Confirm it's running:
```bash
sudo systemctl status prometheus
```
Prometheus UI: http://localhost:9090

### 2. node_exporter
```bash
sudo apt install prometheus-node-exporter
```
It exposes metrics on port `9100`. Make sure Prometheus scrapes it — you should have a job like this in `/etc/prometheus/prometheus.yml`:
```yaml
  - job_name: node
    static_configs:
      - targets: ['localhost:9100']
```

### 3. Grafana (optional, for dashboards)
```bash
sudo apt install -y apt-transport-https software-properties-common
# add Grafana APT repo per official docs, then:
sudo apt update
sudo apt install grafana
sudo systemctl enable --now grafana-server
```
Grafana UI: http://localhost:3000 (default login `admin` / `admin` — change it immediately).

> Note: if you rename the admin user and forget it, reset with:
> ```bash
> grafana-cli admin reset-admin-password NEW_PASSWORD
> ```

---

## Usage

1. Make sure Prometheus has been **running and scraping for the whole period you want to capture**. The script only exports what Prometheus already stored — if Prometheus was down during your test, that data does not exist.
2. Set the time window at the top of the script:
   ```python
   HOURS = 4        # how many hours back to export
   STEP  = "10s"    # sample resolution
   ```
3. Run:
   ```bash
   python3 export_metrics_fixed.py
   ```
4. Output: `prometheus_metrics_YYYYMMDD_HHMMSS.xlsx` in the current directory.

---

## Configuration

| Variable | Default | Meaning |
|----------|---------|---------|
| `PROMETHEUS` | `http://localhost:9090` | Prometheus base URL |
| `HOURS` | `4` | How far back to export |
| `STEP` | `10s` | Sample interval (smaller = more rows, more detail) |
| `SECTIONS` | dict | Prefix → sheet mapping; anything unmatched goes to **Other** |

**A note on `STEP` and file size:** at `10s` over several hours the **Other** sheet alone can exceed the Excel row limit. The script handles this by splitting into `Other_2`, etc. If you want smaller files, raise `STEP` to `30s`.

---

## GPU monitoring (NVIDIA) — optional add-on

> There are two versions of the exporter script: a base one (node metrics only) and this GPU-enabled one. Use the GPU version only if you have an NVIDIA card and the exporter installed.

### Install nvidia_gpu_exporter
```bash
cd ~
VERSION=1.4.1
wget https://github.com/utkuozdemir/nvidia_gpu_exporter/releases/download/v${VERSION}/nvidia_gpu_exporter_${VERSION}_linux_x86_64.tar.gz
tar -xvzf nvidia_gpu_exporter_${VERSION}_linux_x86_64.tar.gz
sudo mv nvidia_gpu_exporter /usr/bin/
```

### Run it
```bash
/usr/bin/nvidia_gpu_exporter &
```
It listens on port `9835`.

### Tell Prometheus to scrape it
Add to the end of `/etc/prometheus/prometheus.yml` (mind the indentation — 2 spaces before `-`, matching the other jobs):
```yaml
  - job_name: nvidia_gpu
    static_configs:
      - targets: ['localhost:9835']
```
Then restart Prometheus:
```bash
sudo systemctl restart prometheus
```

### Verify GPU metrics are flowing
```bash
curl -s http://localhost:9835/metrics | grep temperature_gpu
```
You should see a line like `nvidia_smi_temperature_gpu 42`.

Check Prometheus picked up the target:
```bash
curl -s http://localhost:9090/api/v1/targets | grep nvidia_gpu
```

### Autostart across reboots (optional)
The exporter does **not** survive a reboot when launched with `&`. To make it permanent, create a systemd service:
```bash
sudo nano /etc/systemd/system/nvidia_gpu_exporter.service
```
```ini
[Unit]
Description=NVIDIA GPU Exporter
After=network.target

[Service]
ExecStart=/usr/bin/nvidia_gpu_exporter
Restart=always
User=root

[Install]
WantedBy=multi-user.target
```
```bash
sudo systemctl daemon-reload
sudo systemctl enable --now nvidia_gpu_exporter
sudo systemctl status nvidia_gpu_exporter
```

### What GPU data you get
The GPU sheet includes, among others:
- `nvidia_smi_temperature_gpu` — GPU core temperature
- `nvidia_smi_memory_used_bytes` / `memory_total_bytes` — VRAM usage
- `nvidia_smi_pstate` — power state (P0 = full load, P8 = idle)

> **Legacy-driver caveat:** on older cards (e.g. Maxwell GTX 750 on the `nvidia-470` legacy driver) `nvidia-smi` may report `utilization.gpu` as *not supported*, so GPU-utilization % may be missing even when temperature and VRAM work fine. Try `nvidia-settings -t -q GPUUtilization` as an alternative source.

---

## Troubleshooting

**`ConnectionError ... port=9090 Connection refused`**
Prometheus isn't running or crashed. Check:
```bash
sudo systemctl status prometheus
```
A common cause is a bad edit to `prometheus.yml` (YAML is indentation-sensitive). After any edit, restart and check the status — `INVALIDARGUMENT` / `Error loading config` means a syntax/indentation error.

**`ValueError: Row numbers must be between 1 and 1048576`**
A sheet exceeded Excel's row cap. This version already splits large sheets automatically; if you hit it, make sure you're running the fixed script (the one with `MAX_ROWS`).

**GPU sheet is empty / missing**
- The exporter wasn't running during the test window, or
- Prometheus wasn't scraping it, or
- The metric name prefix isn't in `SECTIONS`.
Re-check the three verify commands in the GPU section above **before** starting a long test.

**Exporter: `bind: address already in use` on :9835**
An old instance is still running. Find and stop it:
```bash
sudo lsof -i :9835
kill <PID>
```

**`curl` shows an HTML page instead of metrics**
You hit `/netrics` instead of `/metrics` (easy typo). Use `-s` to hide the progress bar:
```bash
curl -s http://localhost:9835/metrics | grep temperature_gpu
```

---

## Important workflow note

Prometheus is the source of truth. **Start Prometheus (and the GPU exporter, if used) before your test and keep them running for its entire duration.** The exporter launched with `&` dies on reboot — do not reboot between the test and running the export script, or the data window will be incomplete.

---

## License
Personal project — use freely.
