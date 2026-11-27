# Hikvision-ISAPI-NVR-Audit-Excel-Exporter
Hikvision-ISAPI-NVR-Audit-Excel-Exporter

# Hikvision / ISAPI NVR Audit & Health Exporter

A standalone, multi-threaded PowerShell utility that concurrently queries one or more ISAPI-compatible Network Video Recorders (NVRs) and exports formatted Excel (`.xlsx`) audit reports directly to your Desktop.

Built for IT and physical security administrators who need rapid visibility into NVR hardware metadata, camera online/offline statuses, video stream configurations, recording retention spans, HDD health, and user accounts without relying on heavy VMS client software.

---

## Key Features

* **Parallel Execution:** Processes multiple NVRs simultaneously using PowerShell background jobs (`Start-Job`), generating a separate timestamped workbook per device.
* **Zero-Persistence Credential Hygiene:** Prompts interactively for credentials at runtime, validates authentication before launching workers, holds passwords strictly in RAM as `SecureString`, and purges them from memory upon completion.
* **Smart Recording Retention Calculation:** Queries ISAPI content search endpoints to identify the oldest and newest recorded footage per camera channel and calculates exact retention days.
* **Formatted Multi-Sheet Excel Reports:** Automatically installs and uses the `ImportExcel` module (no Microsoft Excel installation required on the host machine) to produce clean, border-styled worksheets with conditional green/red status highlighting.
* **User Attribution:** Logs the full name of the administrator generating the report alongside export timestamps for audit trails.

---

## Report Structure

Each NVR generates a file named `NVR_login_<IP>_<Timestamp>.xlsx` on the user's Desktop containing four worksheets:

1. **`NVR_Details`**: Executive summary banner (online/offline camera counts, total storage usage) plus device name, model, serial number, firmware version, MAC address, network configuration, uptime, and NTP settings.
2. **`Cameras`**: Channel-by-channel inventory including camera name, IP address, protocol, online/offline status, camera model/serial/firmware (via proxy/video input fallback queries), resolution, codec, FPS, bitrate, oldest/newest recording timestamps, and total retention days.
3. **`NVR_Storage`**: HDD slot ID, status, drive type, total capacity, used space, and free space in GB.
4. **`Users`**: Configured NVR user accounts, privilege levels, and IP/MAC binding restrictions.

---

## Prerequisites

* **OS:** Windows 10 / 11 or Windows Server (PowerShell 5.1 or PowerShell 7+)
* **Network:** HTTP/HTTPS access to the target NVR(s) with ISAPI enabled
* **Permissions:** An NVR account with administrative/query access to `/ISAPI/System`, `/ISAPI/ContentMgmt`, `/ISAPI/Streaming`, and `/ISAPI/Security` endpoints
* **Dependencies:** The `ImportExcel` PowerShell module (automatically installed for the current user if missing)

---

## Usage

1. Clone or download this repository:
   ```powershell
   git clone https://github.com/xitizbasnet/Hikvision-ISAPI-NVR-Audit-Excel-Exporter.git
   ```
  ```powershell
      cd hikvision-isapi-nvr-audit
   ```



1. Run the script in PowerShell:
   ```powershell
   .\Export-NvrAuditReport.ps1
   ```

2. Follow the interactive prompts:
   * **Full Name:** Enter your name for report attribution.
   * **NVR IP(s) or Domain(s):** Enter a single address or a comma-separated list (e.g., `192.168.1.10, 192.168.1.11, nvr.example.com`).
   * **Credentials:** Enter your NVR login username and password.

3. Once complete, the generated `.xlsx` files will save to your Desktop and open automatically.

---

## Security Notes

* **Self-Signed Certificates:** The script bypasses SSL certificate validation to support NVRs using default self-signed certificates over HTTPS.
* **Credential Safety:** No usernames, passwords, or IP lists are written to disk or hardcoded in the script. All sensitive variables are disposed of in the `finally` block.
* **Shared Credentials:** Currently assumes the same username and password across all comma-separated NVRs entered in a single run.

---

## License

This project is licensed under the MIT License.
