# 🎭 ppt-orchestrator

> **A zero-dependency, robust web-based PowerPoint orchestration tool built for real-world stage managers and event staff.**

`ppt-orchestrator` enables seamless, zero-downtime transitions between multiple speakers' slide decks using just a smartphone or tablet. It requires **no third-party software installation** (like Node.js or Python) on the host PC, making it the perfect solution for strict corporate environments where installing external software is prohibited.

---

## 💡 Why This Exists (Solving Real-World Challenges)

Running a multi-speaker event often comes with severe operational headaches. This tool was engineered specifically to solve them:

* 🚫 **The "No Install" Constraint:** Corporate PCs often block third-party software. **Solution:** Built entirely on native Windows tools (PowerShell & Batch). If you have Windows and PowerPoint, it just works.
* 📶 **Unstable Venue Wi-Fi:** Mobile devices go to sleep, or venue Wi-Fi drops momentarily. **Solution:** Engineered with a robust polling architecture (V7.3+). The server safely ignores broken pipes and client disconnects. The presentation will *never* crash due to a network drop.
* 📺 **Ugly Desktop Transitions:** Dragging windows or showing the file explorer to the audience looks unprofessional. **Solution:** One-tap seamless transitions. The next slide deck opens directly in full-screen mode the moment the previous one ends.
* ⚠️ **Human Error Under Pressure:** Clicking the wrong file or repeating a speaker during a fast-paced event. **Solution:** Smart queue management. Finished presentations are automatically moved to a `finish/` directory.

---

## ✨ Key Features

- **📱 Mobile Web Remote:** Control presentations directly from your phone's browser.
- **🔄 Seamless Transitions:** Instant, full-screen switching between `.ppt` / `.pptx` files.
- **🗂️ Auto-Queue Management:** Automatically detects presentation files and sorts them into "Pending" and "Completed" lists.
- **🛡️ Auto-Configuration:** The included Batch file automatically handles Administrator elevation, Windows Firewall rules, and URLACL network bindings.

---

## 📂 Directory Structure

Place your PowerPoint files in the same directory as the scripts. The tool will automatically manage them.

```text
ppt-orchestrator/
│
├── Start-Presenter.bat        # 🚀 The Launcher (Handles Elevation & Network config)
├── Invoke-PPTController.ps1   # ⚙️ The Core Server & Logic
├── README.md                  # 📖 Documentation
│
├── 01_Opening_Remarks.pptx    # ⬅️ Drop your presentation files here
├── 02_Keynote_Speech.pptx     # ⬅️ (Files are sorted alphabetically)
│
└── finish/                    # 📁 Auto-generated folder
    └── 00_Test_Slide.pptx     # ⬅️ Finished files are automatically moved here

```

---

## 🚀 How to Use

### Prerequisites

1. A Windows PC with **Microsoft PowerPoint** installed.
2. The PC and your mobile device must be on the **same network** (Wi-Fi or LAN).

### Step-by-Step Guide

1. **Prepare the Files**
Place all your `.ppt` or `.pptx` files in the root folder containing the scripts. (Tip: Prefix file names with numbers like `01_...`, `02_...` for guaranteed order).
2. **Launch the Server**
Double-click `Start-Presenter.bat`.
* *Note: It will prompt for Administrator privileges (UAC). This is required to temporarily configure the Windows Firewall and `http.sys` to allow web traffic.*


3. **Connect Your Remote**
The console window will display a URL (e.g., `http://192.168.x.x:8090/`). Open this URL in the web browser of your smartphone or tablet.
4. **Control the Show**
* **Lobby Screen:** Select a specific file from the queue or simply press **"Start"**.
* **Now Presenting:** The slide will open full-screen on the PC. You can monitor the status on your phone.
* **Post-Presentation:** When a slide deck ends (or is manually stopped), the file is moved to the `finish/` folder. Your phone will prompt you to start the "Next" slide, Retry, or return to the Lobby.


5. **Clean Shutdown**
Click **"Exit System"** on your mobile device, or press `Q` in the PC console. The script will safely close PowerPoint, remove the temporary Firewall rules, and clean up network bindings.

---

## 🛠️ Advanced Operations (PC Console)

If the Wi-Fi completely fails and you lose access to the web remote, the host PC operator can still fully control the flow using the keyboard in the console window:

* `[Enter]` or `[S]`: Start / Next Slide
* `[1] - [9]`: Jump to a specific file in the queue
* `[N] / [P]`: Navigate pages (if you have many files)
* `[Q]` or `[ESC]`: Safely shut down the system

---

## ❓ Troubleshooting

**Q: The web page isn't loading on my phone.**

> Ensure your phone and the host PC are connected to the exactly same Wi-Fi network. Also, verify that the PC's network profile is set to "Private" (though the batch file attempts to allow the port across all profiles).

**Q: I get a "Port in use" error.**

> Port `8090` might be used by another application. You can change the port by editing `set "WEB_PORT=8090"` in the `.bat` file and `[int]$WebPort = 8090` in the `.ps1` file.

**Q: Does it support clickers / presentation remotes?**

> Yes! The speaker can use a standard USB clicker to advance their slides while they are speaking. `ppt-orchestrator` simply manages the "switching" between files behind the scenes.
