[README.md](https://github.com/user-attachments/files/34185352/README.md)
# Group Contribution Tracker 2026

A green-themed Flask web application for tracking monthly group contributions.

## Features

- **Login page** with colourful green theme
- **Dashboard** showing all members, monthly payments, totals and balances
- Colour coding:
  - Green cells / names → members who left on 7 July 2026 (no further penalties)
  - Red / missed cells → failed to pay by the 5th
  - Yellow / late → paid with KES 300 penalty (shown as 2,300)
  - Green paid → on-time KES 2,000
- **Individual member statements** with downloadable TXT
- **Group statement download** (CSV and full TXT)
- Rules enforced in the UI:
  - Monthly contribution: **KES 2,000** due by the **5th** of every month
  - Late penalty: **KES 300**

## Demo login

| Username   | Password   |
|------------|------------|
| admin      | admin123   |
| treasurer  | green2026  |

## How to run

```bash
cd contribution_app
pip install -r requirements.txt
python app.py
```

Then open http://127.0.0.1:5000 in your browser.

## Project structure

```
contribution_app/
├── app.py                 # Main Flask application + data
├── requirements.txt
├── README.md
├── static/
│   ├── css/style.css      # Green theme styles
│   └── js/main.js
└── templates/
    ├── base.html
    ├── login.html
    ├── dashboard.html
    └── member.html
```
