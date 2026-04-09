═══════════════════════════════════════════════════════
  IGNOU RC-71 · Filter & Award List Generator  v2
  100% Offline · No Internet Required
═══════════════════════════════════════════════════════

REQUIREMENTS (one-time setup):
  Python 3.8+  →  https://python.org/downloads
  Open Terminal and run:
    pip install flask python-docx openpyxl

HOW TO RUN:
  Windows : Double-click  START.bat
  Mac/Linux: Run  ./start.sh  or  python3 app.py
  Then open browser →  http://localhost:5050

HOW TO USE:
  1. Upload your Excel / CSV file
  2. Pick TEE Session: June or December + Year
  3. Click course pills to filter (multi-select supported)
  4. The summary bar shows how many pages will be generated
  5. Click "Export IGNOU Award List (.docx)"

OUTPUT:
  • One page per course — each page is a complete IGNOU
    Award/Grade List form with correct headers
  • If 5 courses selected → 5-page Word document
  • Each page sorted by Enrollment No. ascending
  • Session label auto-filled: "Jun 2024" or "Dec 2024" etc.
  • Min 25 rows per page (blank rows keep the form size)
