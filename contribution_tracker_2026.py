#!/usr/bin/env python3
"""
==============================================================================
  GROUP CONTRIBUTION TRACKER 2026
  Complete single-file Flask web application
==============================================================================

Features:
  - Colourful green-themed LOGIN page (username + password)
  - Dashboard with all members, monthly payments, balances
  - Individual member statements (view + download)
  - Full group statement download (CSV + TXT)
  - Red cells = missed payment by 5th (penalty KES 300)
  - Green members = left 7 July 2026 (no further penalties)
  - Monthly contribution: KES 2,000 due by 5th of every month

Demo logins:
  admin / admin123
  treasurer / green2026

How to run:
  pip install flask
  python contribution_tracker_2026.py
  Then open http://127.0.0.1:5000

==============================================================================
"""

from flask import (
    Flask, request, redirect, url_for, session, flash,
    send_file, make_response, render_template_string
)
from functools import wraps
from io import BytesIO, StringIO
from datetime import datetime
import csv

app = Flask(__name__)
app.secret_key = "contribution-tracker-2026-green-theme-secret-key-change-in-prod"

# ---------------------------------------------------------------------------
# Demo credentials
# ---------------------------------------------------------------------------
USERS = {
    "admin": "admin123",
    "treasurer": "green2026",
}

# ---------------------------------------------------------------------------
# Data from the Excel file
# ---------------------------------------------------------------------------
MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUNE", "JULY", "AUG", "SEPT", "OCT", "NOV", "DEC"]

MEMBERS = [
    {"no": 1, "name": "JOHNSTONE MBUVI", "prev_total": 36000,
     "payments": {"JAN": None, "FEB": None, "MAR": None, "APR": None, "MAY": None, "JUNE": None, "JULY": None, "AUG": None, "SEPT": None, "OCT": None, "NOV": None, "DEC": None},
     "total": 36000, "balance": 32200, "left": False},
    {"no": 2, "name": "FAITH ACHIENG", "prev_total": 51200,
     "payments": {"JAN": 2000, "FEB": 2000, "MAR": None, "APR": None, "MAY": 2000, "JUNE": None, "JULY": None, "AUG": None, "SEPT": None, "OCT": None, "NOV": None, "DEC": None},
     "total": 57200, "balance": 16100, "left": False},
    {"no": 3, "name": "BRAYAN MBOGO", "prev_total": 50300,
     "payments": {"JAN": 2000, "FEB": 2000, "MAR": 2000, "APR": 2000, "MAY": 2000, "JUNE": 2000, "JULY": 2000, "AUG": 2000, "SEPT": 2300, "OCT": 2000, "NOV": None, "DEC": None},
     "total": 70600, "balance": 2300, "left": False},
    {"no": 4, "name": "DANIEL GITHANGA", "prev_total": 44500,
     "payments": {"JAN": None, "FEB": None, "MAR": None, "APR": None, "MAY": None, "JUNE": None, "JULY": None, "AUG": None, "SEPT": None, "OCT": None, "NOV": None, "DEC": None},
     "total": 44500, "balance": 35900, "left": False},
    {"no": 5, "name": "ALPHONCE OCHIENG", "prev_total": 19500,
     "payments": {"JAN": None, "FEB": None, "MAR": None, "APR": None, "MAY": None, "JUNE": None, "JULY": None, "AUG": None, "SEPT": None, "OCT": None, "NOV": None, "DEC": None},
     "total": 19500, "balance": 51500, "left": True},
    {"no": 6, "name": "FREDRICK MAINA", "prev_total": 55000,
     "payments": {"JAN": None, "FEB": None, "MAR": None, "APR": None, "MAY": 2000, "JUNE": 2000, "JULY": None, "AUG": None, "SEPT": None, "OCT": None, "NOV": None, "DEC": None},
     "total": 59000, "balance": 16100, "left": False},
    {"no": 7, "name": "MARGARET KIRIGO", "prev_total": 25000,
     "payments": {"JAN": None, "FEB": None, "MAR": None, "APR": None, "MAY": None, "JUNE": None, "JULY": None, "AUG": None, "SEPT": None, "OCT": None, "NOV": None, "DEC": None},
     "total": 25000, "balance": 48900, "left": False},
    {"no": 8, "name": "FRANKLINE OKOTI", "prev_total": 34000,
     "payments": {"JAN": None, "FEB": None, "MAR": None, "APR": None, "MAY": None, "JUNE": None, "JULY": None, "AUG": None, "SEPT": None, "OCT": None, "NOV": None, "DEC": None},
     "total": 34000, "balance": 30000, "left": True},
    {"no": 9, "name": "EMMANUEL BOSIRE", "prev_total": 16000,
     "payments": {"JAN": None, "FEB": None, "MAR": None, "APR": None, "MAY": None, "JUNE": None, "JULY": None, "AUG": None, "SEPT": None, "OCT": None, "NOV": None, "DEC": None},
     "total": 16000, "balance": 49800, "left": False},
    {"no": 10, "name": "THOMAS ODHIAMBO", "prev_total": 52000,
     "payments": {"JAN": 2000, "FEB": 2000, "MAR": None, "APR": None, "MAY": None, "JUNE": None, "JULY": None, "AUG": None, "SEPT": None, "OCT": None, "NOV": None, "DEC": None},
     "total": 56000, "balance": 9800, "left": True},
    {"no": 11, "name": "JOHN NJUGUNA", "prev_total": 54700,
     "payments": {"JAN": 2000, "FEB": 2300, "MAR": 2300, "APR": 2300, "MAY": 2300, "JUNE": 2300, "JULY": 2300, "AUG": 2000, "SEPT": None, "OCT": None, "NOV": None, "DEC": None},
     "total": 72500, "balance": 2300, "left": False},
    {"no": 12, "name": "STEPHEN YIEMBE", "prev_total": 50000,
     "payments": {"JAN": 2000, "FEB": 2000, "MAR": 2000, "APR": 2000, "MAY": 2000, "JUNE": 2000, "JULY": 2300, "AUG": 2300, "SEPT": 2300, "OCT": 1750, "NOV": None, "DEC": None},
     "total": 70650, "balance": 0, "left": False},
    {"no": 13, "name": "HENRY OGADA", "prev_total": 52330,
     "payments": {"JAN": 2000, "FEB": 2000, "MAR": None, "APR": 2000, "MAY": 2000, "JUNE": None, "JULY": None, "AUG": None, "SEPT": None, "OCT": None, "NOV": None, "DEC": None},
     "total": 60330, "balance": 10500, "left": False},
    {"no": 14, "name": "GILBERT CHERUIYOT", "prev_total": 56600,
     "payments": {"JAN": 2300, "FEB": 2300, "MAR": 2300, "APR": 2300, "MAY": 2300, "JUNE": 2300, "JULY": 2300, "AUG": 2300, "SEPT": 2300, "OCT": 2000, "NOV": None, "DEC": None},
     "total": 79300, "balance": 0, "left": False},
]

DIVIDENDS = 49530
CUMULATIVE_PREV = 646660
CUMULATIVE_TOTAL = 736310
CUMULATIVE_BALANCE = 305400
MONTHLY_TARGET = 2000
PENALTY = 300

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("logged_in"):
            flash("Please log in to continue.", "warning")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated


def get_member(no):
    for m in MEMBERS:
        if m["no"] == no:
            return m
    return None


def payment_status(member, month):
    val = member["payments"].get(month)
    if member["left"] and month in ("AUG", "SEPT", "OCT", "NOV", "DEC"):
        return "left", None
    if val is None:
        return "missed", None
    if val >= MONTHLY_TARGET + PENALTY:
        return "late", val
    if val == MONTHLY_TARGET:
        return "paid", val
    if val > 0:
        return "partial", val
    return "missed", None


def member_statement_lines(member):
    lines = []
    lines.append(f"MEMBER STATEMENT – {member['name']}")
    lines.append(f"Member No: {member['no']}")
    lines.append(f"Previous years total (23/24/25): KES {member['prev_total']:,}")
    lines.append("-" * 50)
    lines.append("2026 Monthly Contributions (KES 2,000 due by 5th of each month)")
    lines.append(f"Late payment penalty: KES {PENALTY}")
    if member["left"]:
        lines.append("NOTE: Member left on 7 July 2026 – no penalties after departure.")
    lines.append("-" * 50)
    year_contrib = 0
    for month in MONTHS:
        status, val = payment_status(member, month)
        if status == "left":
            lines.append(f"  {month:6}: LEFT (no longer contributing)")
        elif status == "paid":
            lines.append(f"  {month:6}: {val:,}  (on time)")
            year_contrib += val
        elif status == "late":
            lines.append(f"  {month:6}: {val:,}  (includes KES {PENALTY} penalty)")
            year_contrib += val
        elif status == "partial":
            lines.append(f"  {month:6}: {val:,}  (partial payment)")
            year_contrib += val
        else:
            lines.append(f"  {month:6}: MISSED / not paid")
    lines.append("-" * 50)
    lines.append(f"2026 contributions total : KES {year_contrib:,}")
    lines.append(f"Grand total contributed  : KES {member['total']:,}")
    lines.append(f"BALANCE TO BE PAID       : KES {member['balance']:,}")
    lines.append("-" * 50)
    lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    return lines


# ---------------------------------------------------------------------------
# CSS (green theme)
# ---------------------------------------------------------------------------
CSS = """
:root {
    --green-900:#0d3b2e; --green-800:#145c45; --green-700:#1a7a5a; --green-600:#228b5a;
    --green-500:#2eaa6e; --green-400:#4cbe86; --green-200:#b8e8d0; --green-100:#e6f7ef; --green-50:#f0faf5;
    --lime:#a8e063; --teal:#0d9488; --forest:#166534; --emerald:#059669; --orange:#ea580c;
    --red:#dc2626; --amber:#d97706; --white:#fff; --gray-50:#f9fafb; --gray-100:#f3f4f6;
    --gray-200:#e5e7eb; --gray-600:#4b5563; --gray-800:#1f2937;
    --shadow:0 4px 20px rgba(13,59,46,0.12); --radius:12px;
}
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',system-ui,sans-serif;background:linear-gradient(160deg,var(--green-50) 0%,#e8f5e9 40%,#f1f8e9 100%);color:var(--gray-800);min-height:100vh;display:flex;flex-direction:column}
.navbar{background:linear-gradient(135deg,var(--green-900),var(--green-700));color:#fff;padding:.85rem 1.75rem;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:.75rem;box-shadow:0 2px 12px rgba(0,0,0,.2);position:sticky;top:0;z-index:100}
.nav-brand{display:flex;align-items:center;gap:.6rem;font-size:1.2rem;font-weight:700}
.nav-brand i{color:var(--lime);font-size:1.4rem}
.nav-links{display:flex;align-items:center;gap:.5rem;flex-wrap:wrap}
.nav-links a{color:rgba(255,255,255,.9);text-decoration:none;padding:.4rem .85rem;border-radius:8px;font-size:.9rem;transition:background .2s}
.nav-links a:hover,.nav-links a.active{background:rgba(255,255,255,.18);color:#fff}
.user-badge{background:rgba(255,255,255,.15);padding:.35rem .75rem;border-radius:20px;font-size:.85rem}
.logout-btn{background:rgba(220,38,38,.85)!important}
.container{max-width:1400px;margin:0 auto;padding:1.5rem 1.25rem 3rem;flex:1;width:100%}
.flash{padding:.85rem 1.1rem;border-radius:8px;margin-bottom:.5rem;display:flex;justify-content:space-between;align-items:center;font-size:.95rem}
.flash-success{background:#d1fae5;color:#065f46;border-left:4px solid var(--emerald)}
.flash-danger{background:#fee2e2;color:#991b1b;border-left:4px solid var(--red)}
.flash-warning{background:#fef3c7;color:#92400e;border-left:4px solid var(--amber)}
.flash-info{background:#e0f2fe;color:#075985;border-left:4px solid #0284c7}
.flash-close{background:none;border:none;font-size:1.3rem;cursor:pointer;color:inherit;opacity:.7}
.page-header{margin-bottom:1.5rem}
.page-header h1{font-size:1.75rem;color:var(--green-900);display:flex;align-items:center;gap:.5rem;flex-wrap:wrap}
.page-sub{color:var(--gray-600);margin-top:.25rem;font-size:.95rem}
.back-link{display:inline-flex;align-items:center;gap:.35rem;color:var(--green-700);text-decoration:none;font-size:.9rem;margin-bottom:.5rem}
.stats-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:1rem;margin-bottom:1.5rem}
.stat-card{background:#fff;border-radius:var(--radius);padding:1.1rem 1.2rem;display:flex;align-items:center;gap:.9rem;box-shadow:var(--shadow);border-top:4px solid transparent;transition:transform .2s}
.stat-card:hover{transform:translateY(-3px)}
.stat-card.green{border-color:var(--green-600)}.stat-card.teal{border-color:var(--teal)}.stat-card.lime{border-color:#84cc16}
.stat-card.forest{border-color:var(--forest)}.stat-card.orange{border-color:var(--orange)}.stat-card.emerald{border-color:var(--emerald)}
.stat-icon{width:48px;height:48px;border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:1.3rem;color:#fff}
.stat-card.green .stat-icon{background:var(--green-600)}.stat-card.teal .stat-icon{background:var(--teal)}
.stat-card.lime .stat-icon{background:#84cc16}.stat-card.forest .stat-icon{background:var(--forest)}
.stat-card.orange .stat-icon{background:var(--orange)}.stat-card.emerald .stat-icon{background:var(--emerald)}
.stat-value{display:block;font-size:1.25rem;font-weight:700;color:var(--green-900)}
.stat-label{font-size:.8rem;color:var(--gray-600)}
.legend-box{background:#fff;border-radius:var(--radius);padding:1rem 1.25rem;margin-bottom:1.25rem;box-shadow:var(--shadow)}
.legend-box h3{font-size:1rem;color:var(--green-800);margin-bottom:.6rem}
.legend-items{display:flex;flex-wrap:wrap;gap:.75rem 1.25rem}
.leg{display:inline-flex;align-items:center;gap:.35rem;font-size:.85rem;padding:.25rem .6rem;border-radius:6px}
.leg.paid{background:#d1fae5;color:#065f46}.leg.late{background:#fef3c7;color:#92400e}
.leg.missed{background:#fee2e2;color:#991b1b}.leg.left{background:#dcfce7;color:#166534}.leg.partial{background:#e0e7ff;color:#3730a3}
.download-bar{display:flex;flex-wrap:wrap;gap:.75rem;margin-bottom:1.25rem}
.btn{display:inline-flex;align-items:center;gap:.4rem;padding:.65rem 1.2rem;border-radius:8px;font-size:.9rem;font-weight:600;text-decoration:none;border:none;cursor:pointer;transition:background .2s,transform .15s}
.btn:hover{transform:translateY(-1px)}
.btn-download{background:linear-gradient(135deg,var(--green-700),var(--green-500));color:#fff;box-shadow:0 2px 8px rgba(34,139,90,.35)}
.btn-download:hover{background:linear-gradient(135deg,var(--green-800),var(--green-600))}
.btn-download-alt{background:#fff;color:var(--green-800);border:2px solid var(--green-500)}
.btn-download-alt:hover{background:var(--green-100)}
.btn-sm{padding:.35rem .7rem;font-size:.8rem;border-radius:6px}
.btn-view{background:var(--green-100);color:var(--green-800);text-decoration:none}
.btn-view:hover{background:var(--green-200)}
.table-wrapper{overflow-x:auto;background:#fff;border-radius:var(--radius);box-shadow:var(--shadow);margin-bottom:1.5rem}
.members-table{width:100%;border-collapse:collapse;font-size:.82rem;min-width:1100px}
.members-table th{background:linear-gradient(135deg,var(--green-800),var(--green-700));color:#fff;padding:.7rem .45rem;text-align:center;font-weight:600;white-space:nowrap}
.members-table th:first-child,.members-table th:nth-child(2){text-align:left;padding-left:.75rem}
.members-table td{padding:.55rem .4rem;text-align:center;border-bottom:1px solid var(--gray-100)}
.members-table td:first-child,.members-table td:nth-child(2){text-align:left;padding-left:.75rem}
.members-table tbody tr:hover{background:var(--green-50)}
.row-left{background:#f0fdf4!important}
.name-cell{font-weight:600;white-space:nowrap}
.name-left{color:var(--forest)}
.badge-left{display:inline-block;background:var(--lime);color:var(--green-900);font-size:.65rem;font-weight:700;padding:.15rem .4rem;border-radius:4px;margin-left:.35rem}
.num{font-variant-numeric:tabular-nums;white-space:nowrap}
.pay-cell{font-size:.78rem;font-weight:500}
.status-paid{background:#d1fae5!important;color:#065f46}.status-late{background:#fef3c7!important;color:#92400e}
.status-missed{background:#fecaca!important;color:#991b1b}.status-left{background:#dcfce7!important;color:#166534}
.status-partial{background:#e0e7ff!important;color:#3730a3}
.total-col{font-weight:700;color:var(--green-800)}
.balance-col{font-weight:700}.bal-zero{color:var(--emerald)}.bal-high{color:var(--red)}
.totals-row{background:var(--green-100)}.totals-row td{border-top:2px solid var(--green-600);padding:.7rem .4rem}
.notes-box{background:#fff;border-radius:var(--radius);padding:1.25rem 1.5rem;box-shadow:var(--shadow);border-left:5px solid var(--green-500)}
.notes-box h3{color:var(--green-800);margin-bottom:.6rem;font-size:1.05rem}
.notes-box ul{padding-left:1.25rem;color:var(--gray-600);line-height:1.7;font-size:.92rem}
.member-summary-cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:1rem;margin-bottom:1.5rem}
.mscard{background:#fff;border-radius:var(--radius);padding:1.1rem 1.25rem;box-shadow:var(--shadow);border-top:4px solid var(--green-500)}
.mscard.ms-balance{border-color:var(--orange)}.mscard.ms-zero{border-color:var(--emerald)}.mscard.ms-left{border-color:var(--lime)}
.ms-label{display:block;font-size:.8rem;color:var(--gray-600);margin-bottom:.25rem}
.ms-value{font-size:1.3rem;font-weight:700;color:var(--green-900)}
.statement-section{background:#fff;border-radius:var(--radius);padding:1.25rem;box-shadow:var(--shadow);margin-bottom:1.5rem}
.section-header{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:.75rem;margin-bottom:1rem}
.section-header h2{font-size:1.15rem;color:var(--green-800)}
.payment-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(110px,1fr));gap:.75rem}
.pay-card{border-radius:10px;padding:.85rem .5rem;text-align:center;border:2px solid transparent}
.pay-card.status-paid{background:#d1fae5;border-color:#6ee7b7}.pay-card.status-late{background:#fef3c7;border-color:#fcd34d}
.pay-card.status-missed{background:#fee2e2;border-color:#fca5a5}.pay-card.status-left{background:#dcfce7;border-color:#86efac}
.pay-card.status-partial{background:#e0e7ff;border-color:#a5b4fc}
.pay-month{font-size:.75rem;font-weight:700;text-transform:uppercase;letter-spacing:.5px;opacity:.8}
.pay-amount{font-size:1.05rem;font-weight:700;margin:.3rem 0}
.pay-status-label{font-size:.7rem;opacity:.85}
.statement-text-box{background:#fff;border-radius:var(--radius);padding:1.25rem;box-shadow:var(--shadow)}
.statement-text-box h2{font-size:1.15rem;color:var(--green-800);margin-bottom:.75rem}
.statement-pre{background:var(--gray-50);border:1px solid var(--gray-200);border-radius:8px;padding:1.25rem;font-family:Consolas,'Courier New',monospace;font-size:.85rem;line-height:1.55;overflow-x:auto;white-space:pre-wrap;margin-bottom:1rem;color:var(--gray-800)}
.footer{background:var(--green-900);color:rgba(255,255,255,.7);text-align:center;padding:1rem;font-size:.8rem;margin-top:auto}
/* LOGIN PAGE */
.login-page{min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;background:linear-gradient(145deg,#0d3b2e 0%,#145c45 35%,#1a7a5a 65%,#2eaa6e 100%);position:relative;overflow:hidden;padding:1.5rem}
.login-bg{position:absolute;inset:0;pointer-events:none;overflow:hidden}
.leaf{position:absolute;color:rgba(255,255,255,.08);font-size:4rem;animation:float 12s ease-in-out infinite}
.leaf-1{top:8%;left:8%;animation-delay:0s;font-size:5rem}.leaf-2{top:60%;right:10%;animation-delay:3s;font-size:6rem}
.leaf-3{bottom:15%;left:20%;animation-delay:6s;font-size:4.5rem}.leaf-4{top:25%;right:25%;animation-delay:9s;font-size:3.5rem}
@keyframes float{0%,100%{transform:translateY(0) rotate(0)}50%{transform:translateY(-25px) rotate(8deg)}}
.login-card{background:#fff;border-radius:20px;padding:2.5rem 2.25rem;width:100%;max-width:400px;box-shadow:0 20px 60px rgba(0,0,0,.3);position:relative;z-index:1}
.login-header{text-align:center;margin-bottom:1.75rem}
.logo-circle{width:72px;height:72px;border-radius:50%;background:linear-gradient(135deg,var(--green-600),var(--lime));display:flex;align-items:center;justify-content:center;margin:0 auto 1rem;font-size:2rem;color:#fff;box-shadow:0 6px 20px rgba(46,170,110,.4)}
.login-header h1{font-size:1.5rem;color:var(--green-900);margin-bottom:.2rem}
.subtitle{color:var(--gray-600);font-size:.9rem}
.login-form .form-group{margin-bottom:1.15rem}
.login-form label{display:block;font-size:.85rem;font-weight:600;color:var(--green-800);margin-bottom:.35rem}
.login-form input{width:100%;padding:.75rem 1rem;border:2px solid var(--gray-200);border-radius:10px;font-size:.95rem;transition:border-color .2s,box-shadow .2s;outline:none}
.login-form input:focus{border-color:var(--green-500);box-shadow:0 0 0 3px rgba(46,170,110,.2)}
.btn-login{width:100%;padding:.85rem;background:linear-gradient(135deg,var(--green-700),var(--green-500));color:#fff;border:none;border-radius:10px;font-size:1rem;font-weight:700;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:.5rem;margin-top:.5rem;transition:transform .15s,box-shadow .2s;box-shadow:0 4px 15px rgba(34,139,90,.35)}
.btn-login:hover{transform:translateY(-2px);box-shadow:0 6px 20px rgba(34,139,90,.45)}
.login-hint{margin-top:1.5rem;text-align:center;font-size:.8rem;color:var(--gray-600);background:var(--green-50);padding:.75rem;border-radius:8px}
.login-footer{position:relative;z-index:1;margin-top:1.5rem;color:rgba(255,255,255,.7);font-size:.8rem;text-align:center}
@media(max-width:768px){.navbar{flex-direction:column;align-items:flex-start}.stats-grid{grid-template-columns:1fr 1fr}}
@media(max-width:480px){.stats-grid{grid-template-columns:1fr}.login-card{padding:1.75rem 1.25rem}}
"""

# ---------------------------------------------------------------------------
# HTML Templates (embedded)
# ---------------------------------------------------------------------------
LOGIN_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Login – Contribution Tracker 2026</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
<style>{{ css }}</style>
</head>
<body class="login-page">
<div class="login-bg">
  <div class="leaf leaf-1"><i class="fas fa-leaf"></i></div>
  <div class="leaf leaf-2"><i class="fas fa-leaf"></i></div>
  <div class="leaf leaf-3"><i class="fas fa-seedling"></i></div>
  <div class="leaf leaf-4"><i class="fas fa-leaf"></i></div>
</div>
<div class="login-card">
  <div class="login-header">
    <div class="logo-circle"><i class="fas fa-leaf"></i></div>
    <h1>Contribution Tracker</h1>
    <p class="subtitle">2026 Group Savings Portal</p>
  </div>
  {% with messages = get_flashed_messages(with_categories=true) %}
    {% if messages %}{% for cat, msg in messages %}<div class="flash flash-{{ cat }}">{{ msg }}</div>{% endfor %}{% endif %}
  {% endwith %}
  <form method="POST" action="{{ url_for('login') }}" class="login-form">
    <div class="form-group">
      <label for="username"><i class="fas fa-user"></i> Username</label>
      <input type="text" id="username" name="username" placeholder="Enter username" required autofocus>
    </div>
    <div class="form-group">
      <label for="password"><i class="fas fa-lock"></i> Password</label>
      <input type="password" id="password" name="password" placeholder="Enter password" required>
    </div>
    <button type="submit" class="btn-login"><i class="fas fa-sign-in-alt"></i> Sign In</button>
  </form>
  <div class="login-hint">
    <p><strong>Demo accounts:</strong></p>
    <p>admin / admin123 &nbsp;|&nbsp; treasurer / green2026</p>
  </div>
</div>
<footer class="login-footer">Monthly contribution KES 2,000 by 5th &bull; Late penalty KES 300</footer>
</body>
</html>
"""

BASE_START = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{% block title %}Contribution Tracker 2026{% endblock %}</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
<style>{{ css }}</style>
</head>
<body>
{% if session.get('logged_in') %}
<nav class="navbar">
  <div class="nav-brand"><i class="fas fa-leaf"></i><span>Contribution Tracker 2026</span></div>
  <div class="nav-links">
    <a href="{{ url_for('dashboard') }}" class="{% if request.endpoint=='dashboard' %}active{% endif %}"><i class="fas fa-tachometer-alt"></i> Dashboard</a>
    <a href="{{ url_for('download_group_statement') }}"><i class="fas fa-file-csv"></i> Group CSV</a>
    <a href="{{ url_for('download_group_txt') }}"><i class="fas fa-file-alt"></i> Full Statement</a>
    <span class="user-badge"><i class="fas fa-user"></i> {{ session.get('username') }}</span>
    <a href="{{ url_for('logout') }}" class="logout-btn"><i class="fas fa-sign-out-alt"></i> Logout</a>
  </div>
</nav>
{% endif %}
<main class="container">
{% with messages = get_flashed_messages(with_categories=true) %}
  {% if messages %}
  <div class="flash-messages">
    {% for cat, msg in messages %}
    <div class="flash flash-{{ cat }}">{{ msg }}<button class="flash-close" onclick="this.parentElement.remove()">&times;</button></div>
    {% endfor %}
  </div>
  {% endif %}
{% endwith %}
"""

BASE_END = """
</main>
<footer class="footer"><p>Group Contribution Tracker &copy; 2026 &bull; Monthly KES 2,000 by 5th &bull; Penalty KES 300</p></footer>
<script>
document.addEventListener('DOMContentLoaded',function(){
  document.querySelectorAll('.flash').forEach(function(el){
    setTimeout(function(){el.style.transition='opacity .4s';el.style.opacity='0';setTimeout(function(){el.remove()},400)},5000);
  });
});
</script>
</body>
</html>
"""

DASHBOARD_HTML = BASE_START + """
<div class="page-header">
  <h1><i class="fas fa-tachometer-alt"></i> Group Dashboard</h1>
  <p class="page-sub">2026 Contributions &bull; Due by 5th of each month &bull; Penalty KES {{ penalty }}</p>
</div>
<div class="stats-grid">
  <div class="stat-card green"><div class="stat-icon"><i class="fas fa-users"></i></div><div class="stat-info"><span class="stat-value">{{ members|length }}</span><span class="stat-label">Total Members</span></div></div>
  <div class="stat-card teal"><div class="stat-icon"><i class="fas fa-user-check"></i></div><div class="stat-info"><span class="stat-value">{{ active_count }}</span><span class="stat-label">Active Members</span></div></div>
  <div class="stat-card lime"><div class="stat-icon"><i class="fas fa-user-minus"></i></div><div class="stat-info"><span class="stat-value">{{ left_count }}</span><span class="stat-label">Left (7 Jul 2026)</span></div></div>
  <div class="stat-card forest"><div class="stat-icon"><i class="fas fa-coins"></i></div><div class="stat-info"><span class="stat-value">KES {{ "{:,}".format(cumulative_total) }}</span><span class="stat-label">Grand Total</span></div></div>
  <div class="stat-card orange"><div class="stat-icon"><i class="fas fa-balance-scale"></i></div><div class="stat-info"><span class="stat-value">KES {{ "{:,}".format(cumulative_balance) }}</span><span class="stat-label">Outstanding Balance</span></div></div>
  <div class="stat-card emerald"><div class="stat-icon"><i class="fas fa-hand-holding-usd"></i></div><div class="stat-info"><span class="stat-value">KES {{ "{:,}".format(dividends) }}</span><span class="stat-label">Dividends</span></div></div>
</div>
<div class="legend-box">
  <h3><i class="fas fa-info-circle"></i> Legend</h3>
  <div class="legend-items">
    <span class="leg paid"><i class="fas fa-check-circle"></i> Paid on time (KES 2,000)</span>
    <span class="leg late"><i class="fas fa-exclamation-circle"></i> Paid with penalty (KES 2,300)</span>
    <span class="leg missed"><i class="fas fa-times-circle"></i> Missed / not paid (red)</span>
    <span class="leg left"><i class="fas fa-door-open"></i> Left 7 Jul 2026 – no further penalty</span>
    <span class="leg partial"><i class="fas fa-adjust"></i> Partial payment</span>
  </div>
</div>
<div class="download-bar">
  <a href="{{ url_for('download_group_statement') }}" class="btn btn-download"><i class="fas fa-file-csv"></i> Download Group CSV</a>
  <a href="{{ url_for('download_group_txt') }}" class="btn btn-download-alt"><i class="fas fa-file-alt"></i> Download Full Statement (TXT)</a>
</div>
<div class="table-wrapper">
<table class="members-table">
<thead>
<tr>
  <th>No</th><th>Member</th><th>Prev Total</th>
  {% for m in months %}<th class="month-col">{{ m }}</th>{% endfor %}
  <th>Total</th><th>Balance</th><th>Action</th>
</tr>
</thead>
<tbody>
{% for member in members %}
<tr class="{% if member.left %}row-left{% endif %}">
  <td>{{ member.no }}</td>
  <td class="name-cell {% if member.left %}name-left{% endif %}">
    {{ member.name }}{% if member.left %}<span class="badge-left">LEFT</span>{% endif %}
  </td>
  <td class="num">{{ "{:,}".format(member.prev_total) }}</td>
  {% for month in months %}
    {% set st = payment_status(member, month) %}
    <td class="pay-cell status-{{ st[0] }}">
      {% if st[0]=='left' %}—{% elif st[1] is not none %}{{ "{:,}".format(st[1]) }}{% else %}—{% endif %}
    </td>
  {% endfor %}
  <td class="num total-col">{{ "{:,}".format(member.total) }}</td>
  <td class="num balance-col {% if member.balance==0 %}bal-zero{% elif member.balance>30000 %}bal-high{% endif %}">{{ "{:,}".format(member.balance) }}</td>
  <td><a href="{{ url_for('member_detail', no=member.no) }}" class="btn-sm btn-view"><i class="fas fa-eye"></i> Statement</a></td>
</tr>
{% endfor %}
</tbody>
<tfoot>
<tr class="totals-row">
  <td colspan="2"><strong>CUMULATIVE / DIVIDENDS</strong></td>
  <td class="num"><strong>{{ "{:,}".format(cumulative_prev) }}</strong></td>
  <td colspan="12"></td>
  <td class="num"><strong>{{ "{:,}".format(cumulative_total) }}</strong></td>
  <td class="num"><strong>{{ "{:,}".format(cumulative_balance) }}</strong></td>
  <td></td>
</tr>
</tfoot>
</table>
</div>
<div class="notes-box">
  <h3><i class="fas fa-sticky-note"></i> Notes</h3>
  <ul>
    <li>Monthly contribution is <strong>KES 2,000</strong>, payable by the <strong>5th</strong> of every month.</li>
    <li>Failure to pay by the 5th attracts a <strong>KES 300</strong> penalty (shown as 2,300).</li>
    <li>Members highlighted in green left the group on <strong>7 July 2026</strong> and are not subject to any further penalties.</li>
    <li>Click <strong>Statement</strong> next to any member to view and download their individual statement.</li>
    <li>Use the buttons above to download the entire group statement (CSV or full text).</li>
  </ul>
</div>
""" + BASE_END

MEMBER_HTML = BASE_START + """
<div class="page-header">
  <a href="{{ url_for('dashboard') }}" class="back-link"><i class="fas fa-arrow-left"></i> Back to Dashboard</a>
  <h1>{% if member.left %}<span class="badge-left">LEFT</span> {% endif %}{{ member.name }}</h1>
  <p class="page-sub">Member No. {{ member.no }} &bull; Individual Contribution Statement</p>
</div>
<div class="member-summary-cards">
  <div class="mscard"><span class="ms-label">Previous Total (23/24/25)</span><span class="ms-value">KES {{ "{:,}".format(member.prev_total) }}</span></div>
  <div class="mscard"><span class="ms-label">2026 + Prev Grand Total</span><span class="ms-value">KES {{ "{:,}".format(member.total) }}</span></div>
  <div class="mscard {% if member.balance==0 %}ms-zero{% else %}ms-balance{% endif %}"><span class="ms-label">Balance to Pay</span><span class="ms-value">KES {{ "{:,}".format(member.balance) }}</span></div>
  {% if member.left %}<div class="mscard ms-left"><span class="ms-label">Status</span><span class="ms-value">Left 7 July 2026</span></div>{% endif %}
</div>
<div class="statement-section">
  <div class="section-header">
    <h2><i class="fas fa-file-invoice"></i> Payment History 2026</h2>
    <a href="{{ url_for('download_member_statement', no=member.no) }}" class="btn btn-download"><i class="fas fa-download"></i> Download Statement</a>
  </div>
  <div class="payment-grid">
  {% for month in months %}
    {% set st = payment_status(member, month) %}
    <div class="pay-card status-{{ st[0] }}">
      <div class="pay-month">{{ month }}</div>
      <div class="pay-amount">{% if st[0]=='left' %}—{% elif st[1] is not none %}KES {{ "{:,}".format(st[1]) }}{% else %}—{% endif %}</div>
      <div class="pay-status-label">
        {% if st[0]=='paid' %}On time{% elif st[0]=='late' %}With penalty{% elif st[0]=='partial' %}Partial{% elif st[0]=='missed' %}Missed{% elif st[0]=='left' %}Left{% endif %}
      </div>
    </div>
  {% endfor %}
  </div>
</div>
<div class="statement-text-box">
  <h2><i class="fas fa-scroll"></i> Full Statement</h2>
  <pre class="statement-pre">{% for line in statement %}{{ line }}
{% endfor %}</pre>
  <a href="{{ url_for('download_member_statement', no=member.no) }}" class="btn btn-download"><i class="fas fa-download"></i> Download as TXT</a>
</div>
""" + BASE_END


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------
@app.route("/")
def index():
    if session.get("logged_in"):
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("logged_in"):
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        if username in USERS and USERS[username] == password:
            session["logged_in"] = True
            session["username"] = username
            flash(f"Welcome, {username}!", "success")
            return redirect(url_for("dashboard"))
        flash("Invalid username or password.", "danger")
    return render_template_string(LOGIN_HTML, css=CSS)


@app.route("/logout")
def logout():
    session.clear()
    flash("You have been logged out.", "info")
    return redirect(url_for("login"))


@app.route("/dashboard")
@login_required
def dashboard():
    active_members = [m for m in MEMBERS if not m["left"]]
    left_members = [m for m in MEMBERS if m["left"]]
    return render_template_string(
        DASHBOARD_HTML,
        css=CSS,
        members=MEMBERS,
        months=MONTHS,
        dividends=DIVIDENDS,
        cumulative_prev=CUMULATIVE_PREV,
        cumulative_total=CUMULATIVE_TOTAL,
        cumulative_balance=CUMULATIVE_BALANCE,
        active_count=len(active_members),
        left_count=len(left_members),
        monthly_target=MONTHLY_TARGET,
        penalty=PENALTY,
        payment_status=payment_status,
    )


@app.route("/member/<int:no>")
@login_required
def member_detail(no):
    member = get_member(no)
    if not member:
        flash("Member not found.", "danger")
        return redirect(url_for("dashboard"))
    statement = member_statement_lines(member)
    return render_template_string(
        MEMBER_HTML,
        css=CSS,
        member=member,
        months=MONTHS,
        statement=statement,
        monthly_target=MONTHLY_TARGET,
        penalty=PENALTY,
        payment_status=payment_status,
    )


@app.route("/download/member/<int:no>")
@login_required
def download_member_statement(no):
    member = get_member(no)
    if not member:
        flash("Member not found.", "danger")
        return redirect(url_for("dashboard"))
    lines = member_statement_lines(member)
    content = "\n".join(lines)
    buffer = BytesIO(content.encode("utf-8"))
    buffer.seek(0)
    filename = f"statement_{member['name'].replace(' ', '_')}_{member['no']}.txt"
    return send_file(buffer, as_attachment=True, download_name=filename, mimetype="text/plain")


@app.route("/download/group")
@login_required
def download_group_statement():
    output = StringIO()
    writer = csv.writer(output)
    header = ["NO", "MEMBER", "PREV_TOTAL"] + MONTHS + ["TOTAL_2026", "BALANCE", "STATUS"]
    writer.writerow(header)
    for m in MEMBERS:
        row = [m["no"], m["name"], m["prev_total"]]
        for month in MONTHS:
            val = m["payments"].get(month)
            row.append(val if val is not None else "")
        status = "LEFT (7 Jul 2026)" if m["left"] else "ACTIVE"
        row.extend([m["total"], m["balance"], status])
        writer.writerow(row)
    writer.writerow([])
    writer.writerow(["DIVIDENDS", "", DIVIDENDS])
    writer.writerow(["CUMULATIVE TOTAL", "", CUMULATIVE_PREV, "", "", "", "", "", "", "", "", "", "", "", "", CUMULATIVE_TOTAL, CUMULATIVE_BALANCE])
    writer.writerow([])
    writer.writerow(["Notes:"])
    writer.writerow(["Monthly contribution: KES 2,000 due by 5th of every month"])
    writer.writerow(["Penalty for late payment: KES 300"])
    writer.writerow(["Members marked LEFT departed on 7 July 2026 and are not subject to further penalties"])
    writer.writerow([f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"])
    buffer = BytesIO(output.getvalue().encode("utf-8"))
    buffer.seek(0)
    return send_file(
        buffer, as_attachment=True,
        download_name=f"group_contribution_statement_2026_{datetime.now().strftime('%Y%m%d')}.csv",
        mimetype="text/csv",
    )


@app.route("/download/group-txt")
@login_required
def download_group_txt():
    lines = []
    lines.append("=" * 70)
    lines.append("GROUP CONTRIBUTION STATEMENT 2026")
    lines.append("=" * 70)
    lines.append(f"Monthly contribution : KES {MONTHLY_TARGET:,} (due by 5th of each month)")
    lines.append(f"Late payment penalty : KES {PENALTY}")
    lines.append(f"Generated            : {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lines.append("=" * 70)
    lines.append("")
    for m in MEMBERS:
        lines.extend(member_statement_lines(m))
        lines.append("")
        lines.append("")
    lines.append("=" * 70)
    lines.append("GROUP TOTALS")
    lines.append(f"  Previous years cumulative : KES {CUMULATIVE_PREV:,}")
    lines.append(f"  Dividends                 : KES {DIVIDENDS:,}")
    lines.append(f"  2026 grand total          : KES {CUMULATIVE_TOTAL:,}")
    lines.append(f"  Outstanding balance       : KES {CUMULATIVE_BALANCE:,}")
    lines.append("=" * 70)
    content = "\n".join(lines)
    buffer = BytesIO(content.encode("utf-8"))
    buffer.seek(0)
    return send_file(
        buffer, as_attachment=True,
        download_name=f"group_full_statement_2026_{datetime.now().strftime('%Y%m%d')}.txt",
        mimetype="text/plain",
    )


if __name__ == "__main__":
    print("=" * 60)
    print("  Contribution Tracker 2026")
    print("  Open http://127.0.0.1:5000 in your browser")
    print("  Login: admin / admin123  or  treasurer / green2026")
    print("=" * 60)
    app.run(host="0.0.0.0", port=5000, debug=True)
