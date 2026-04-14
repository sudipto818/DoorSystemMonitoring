# Door System Monitoring & Signage

A professional dual-application system designed for managing and displaying office door statuses. The system consists of a **Control Panel** (for the owner) and a **Display App** (for the door-mounted screen), synchronized via a local network and a shared SQLite database.

---

## 📂 Codebase Explanation

### Core Applications
- **`src/control_app.py`**: The main entry point for the owner. It initializes the UI and background workers for network synchronization and Outlook calendar polling.
- **`src/display_app.py`**: The entry point for the display hardware (e.g., a tablet or monitor). It runs in full-screen mode and receives real-time updates from the network or database.
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
    2. Automatically opens your **default Calendar App (e.g., Microsoft Outlook)**.
    3. **Action Required**: Once Outlook pops up with the new event, you must paste the ics link in the space provided and then click **Save & Close** to finalize the meeting and sync it to your cloud calendar.

> [!NOTE]
> **Voice Parser Accuracy**: The voice command system uses fuzzy matching to handle slight transcription errors. However, no voice-to-text system is 100% accurate. Recognition quality depends on your microphone and environment; further logic refinements can be made in `src/voice_command.py` to improve complex intent detection.

---

## 🛠 Building the Executables (.exe)

The project includes `.spec` files to package the Python scripts into standalone `.exe` files using **PyInstaller**.

### Prerequisites
Install PyInstaller:
```powershell
pip install pyinstaller
```

### Build Steps
To build the apps, run the following commands from the **root directory** of the project:

1. **Build the Launcher**:
   ```powershell
   pyinstaller .\build\"Door Sign Launcher.spec" --noconfirm
   ```

2. **Build the Control App**:
   ```powershell
   pyinstaller .\build\control_app.spec --noconfirm
   ```

3. **Build the Display App**:
   ```powershell
   pyinstaller .\build\display_app.spec --noconfirm
   ```

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

## ⚙️ Requirements
Ensure you have the following installed to run from source:
- Python 3.10+
- Dependencies listed in `requirements.txt`:
  ```powershell
  pip install -r requirements.txt
  ```

---
