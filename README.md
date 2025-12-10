🔵 Fortigate Traffic Monitor & Counter Reset Tool
A complete Python-based solution for monitoring daily traffic usage of Fortigate firewall policies and resetting traffic counters every 24 hours for accurate reporting.
This project is designed for organizations that need clear, automated visibility into bandwidth consumption per department or policy.

🚀 Features
🔹 1. Daily Traffic Reporting Script
This script automatically:
Reads policy names via Fortigate API
Fetches actual traffic usage (bytes) via SSH diagnose commands
Converts raw data into readable values (GB)
Categorizes policies (Primary/Secondary)
Generates a professional Excel report including:
Color-coded traffic columns
Automatic column sizing
Bar charts and pie charts
Totals & analytics
Stores last usage locally to calculate new changes each day
Runs fully automated using Windows Task Scheduler
Just add your Policy IDs inside the script — everything else is fully automatic.

🔹 2. Counter Reset Script
To keep traffic metrics accurate, a second script is included that:
Connects to Fortigate via SSH
Resets the traffic counter for each configured Policy ID
Ensures the next 24-hour traffic report starts from zero
Can also be scheduled with Windows Task Scheduler
This ensures your daily report always reflects true daily usage, not cumulative data.

🛠️ Requirements
Python 3.x
Fortigate API Token
SSH access to Fortigate
Required Python libraries:
requests
paramiko
openpyxl
Install Python dependencies:
pip install requests paramiko openpyxl

📦 How to Use
1️⃣ Edit config values in the script
Set values like:
FG_IP = "your-firewall-ip"
API_TOKEN = "your-api-token"
POLICY_IDS = [1, 23, 45, ...]

2️⃣ Run manually or schedule automatically
Use Windows Task Scheduler to run:
get data.py → daily report generator
reset policy.py → run 5 min after first file

📊 Output Example
The generated Excel file contains:
Policy ID
Policy Name
Daily Traffic (GB)
Color-coded usage (Low/Medium/High)
Charts for visual analytics
Totals & summary sheet

📁 File Structure
/project-folder
│
├── traffic_report.py       # Main reporting script
├── counter_reset.py        # Counter reset script
└── README.md               # You are here

🧑‍💻
Developed by alireza ahmadi.
If you have questions or suggestions, feel free to open an issue or contact me.
