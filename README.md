# Door System Monitoring & Signage

A professional dual-application system designed for managing and displaying office door statuses. The system consists of a **Control Panel** (for the owner) and a **Display App** (for the door-mounted screen), synchronized via a local network and a shared SQLite database.

---

## 📂 Codebase Explanation

### Core Applications
- **`src/control_app.py`**: The main entry point for the owner. It initializes the UI and background workers for network synchronization and Outlook calendar polling.
- **`src/display_app.py`**: The entry point for the display hardware (e.g., a tablet or monitor). It runs in full-screen mode and receives real-time updates from the network or database.
- For the display hardware, a budget Android tablet is sufficient; an **I Kall N11 Tablet** is a suitable low-cost option.
- **`src/launcher.pyw`**: A lightweight GUI launcher that allows users to easily start either the Control Panel or the Display App without needing a terminal.

### Logic & Helpers
- **`src/db_manager.py`**: Manages the SQLite database (`door_status.db`). It handles the tables for status, visitor messages, and the weekly timetable.
- **`src/network_bridge.py`**: Implements the TCP socket communication layer. It ensures that updates made on the Control Panel are pushed instantly to the Display App over the Wi-Fi network.
- **`src/voice_command.py`**: Integrates AI-powered voice recognition using `faster-whisper`. It allows the owner to set statuses or create meetings simply by speaking.
- **`src/file_store.py`**: Handles persistent configuration data such as IP addresses, ICS URLs, and notification topics in simple text files.
- **`src/optional_deps.py`**: A safety layer that checks for installed dependencies (like Excel or Calendar libraries) and provides graceful fallbacks if they are missing.

### User Interface
- **`src/control_ui.py`**: The complex UI code for the Control Panel, featuring status buttons, visitor logs, and the timetable management interface.
- **`src/display_ui.py`**: The clean, high-visibility UI code for the door display, optimized for readability at a distance.
- **`src/ui_constants.py` & `src/ui_helpers.py`**: Shared design tokens (colors, fonts) and UI components used to maintain a consistent, premium look across both apps.

### Voice Commands & Meeting Flow
The Control Panel features advanced voice recognition (via `faster-whisper`).
- **Commands**: You can set statuses (e.g., "Set status to Busy until 4 PM") or schedule meetings.
- **Meeting Creation**: When you say "Schedule a meeting for [duration]", the app:
    1. Generates a physical `.ics` calendar file in the project folder.
    2. Open your **default Calendar App (e.g., Microsoft Outlook)** and then go to the calendar section, and there you must perform step 3.
    3. **Action Required**: You must paste the ics link in the space provided and then click **Save & Close** to finalize the meeting and sync it to your cloud calendar.

> [!NOTE]
> **Voice Parser Accuracy**: The voice command system uses fuzzy matching to handle slight transcription errors. However, no voice-to-text system is 100% accurate. Recognition quality depends on your microphone and environment; further logic refinements can be made in `src/voice_command.py` to improve complex intent detection.

---

## 🛠 Building the Executables (.exe)

The project includes preconfigured spec files in build/ to package the Python scripts into standalone .exe files using PyInstaller.


### Prerequisites
Install PyInstaller:
```powershell
pip install pyinstaller
```

### Build Steps (Recommended)
Run from the project root folder (DoorSystemMonitoring):

1. Build the Launcher:
   ```powershell
   python -m PyInstaller ".\build\Door Sign Launcher.spec" --noconfirm --distpath .\dist --workpath .\build\build
   ```

2. Build the Control App:
   ```powershell
   python -m PyInstaller .\build\control_app.spec --noconfirm --distpath .\dist --workpath .\build\build
   ```

3. Build the Display App:
   ```powershell
   python -m PyInstaller .\build\display_app.spec --noconfirm --distpath .\dist --workpath .\build\build
   ```

### Where build outputs are created
- Final executables are created in dist/.
- PyInstaller intermediate files are created in build/build/.
- Example executable paths after build:
  - dist/control_app.exe
  - dist/display_app.exe
  - dist/Door Sign Launcher.exe

### Moving executables to another independent location
You can move the built output outside this project folder, but move the full app output (not only a single `.exe`) to avoid missing runtime dependencies.

Recommended:
- Create a release folder outside the project (example: `D:\DoorSignRelease`).
- Copy the required app output from `dist/` into that folder.
- Keep each app's executable with its related files/folders if present.

Important:
- If only the `.exe` is moved and companion files are left behind, the app may fail to start.
- Runtime files such as database/config/temporary files are created relative to where the app runs, so they will be created in the new location after moving.

### How the spec files were created
The spec files were generated from the Python entry scripts and then customized for this project layout.

Original generation method (equivalent):
- pyi-makespec src/control_app.py --name control_app --windowed
- pyi-makespec src/display_app.py --name display_app --windowed
- pyi-makespec src/launcher.pyw --name "Door Sign Launcher" --windowed

Important note on output location:
- Running `pyi-makespec` only creates `.spec` files; it does not create `.exe` files.
- `.exe` output is decided when you run PyInstaller on the `.spec` file.
- If you run from project root (`DoorSystemMonitoring`), output goes to `dist/`.
- If you run from `build/`, output goes to `build/dist/`.
- To avoid confusion, always pass explicit output paths:
  - `--distpath .\\dist`
  - `--workpath .\\build\\build`


---

## 📅 Weekly Timetable Management

The system supports a **Matrix-style** timetable. This allows you to upload a single file that covers your entire week.

### Format Instructions
- **Supported Formats**: `.csv` or `.xlsx` (Excel).
- **Structure**:
    - **Header Row**: The first column must be empty or labeled "Time". The following columns must be the days of the week: `MONDAY`, `TUESDAY`, `WEDNESDAY`, `THURSDAY`, `FRIDAY`.
    - **Data Rows**: 
        - The first column must contain the time range in `HH:MM - HH:MM` format (e.g., `10:00 - 11:30`).
        - Enter the activity name in the cell where the time row meets the day column.

**Example Matrix:**
| Time | MONDAY | TUESDAY | WEDNESDAY |
| :--- | :--- | :--- | :--- |
| 09:00 - 10:30 | Morning Research | Staff Meeting | Office Hours |
| 14:00 - 15:00 | Lab Supervision | | Project Review |

### How to Upload
1. Open the **Control Panel**.
2. Scroll to the **Weekly Weekday Timetable** section.
3. Click **Upload Weekly Matrix**.
4. Select your file. The app will automatically parse and display it.

---

## 📆 Outlook Calendar Integration

The system can automatically set your status to "In a Meeting" based on your Outlook calendar.

### How to Generate your Private ICS Link
1.  Sign in to [Outlook.com](https://outlook.com).
2.  Click the **Settings** (gear icon) in the top right.
3.  Select **View all Outlook settings**.
4.  Navigate to **Calendar** > **Shared calendars**.
5.  Under the **Publish a calendar** section:
    -   Select the calendar you want to share.
    -   Select **Can view all details** for permissions.
    -   Click **Publish**.
6.  Copy the **ICS link** (it usually ends in `.ics`).
7.  Paste this link into the **Outlook Calendar Sync** section in the Control Panel and click **Connect**.

---

## 📱 ntfy Push Notifications

The Control Panel can send instant push notifications to your phone when a new visitor message is received.

### Setup Steps
1. Install the **ntfy** app on your phone (Android/iOS) from PlayStore, or use the web endpoint `https://ntfy.sh/<your-topic>`.
2. Choose a private topic code (example: `my_door_123`) and subscribe to that same topic in the ntfy app.
3. In the **Control Panel** under **Network Display**, enter the topic in **ntfy.sh topic code** and click **Save**.

### How it Works in This App
- The app stores your topic in `ntfy_topic.txt`.
- Notifications are sent for new visitor rows added after startup (so old records are not re-notified).
- Topic input is sanitized to letters, numbers, `_`, and `-`.

### Security Note
- `ntfy.sh` public topics can be guessed. Use a long, hard-to-guess topic code for better privacy.

---

## ⚙️ Requirements
Ensure you have the following installed to run from source:
- Python 3.10+
- Dependencies listed in `requirements.txt`:
  ```powershell
  pip install -r requirements.txt
  ```

---
