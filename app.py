"""
Group Contribution Tracker 2026
Flask Web Application
Monthly contribution: KES 2,000 due by 5th of each month
Penalty for late payment: KES 300
Members in green left on 7 July 2026 and are exempt from further penalties.
"""

from flask import (
    Flask, render_template, request, redirect, url_for,
    session, flash, send_file, make_response
)
from functools import wraps
from io import BytesIO, StringIO
from datetime import datetime
import csv
import os

app = Flask(__name__)
app.secret_key = "contribution-tracker-2026-green-theme-secret-key-change-in-prod"

# ---------------------------------------------------------------------------
# Demo credentials (change these in production)
# ---------------------------------------------------------------------------
USERS = {
    "admin": "admin123",
    "treasurer": "green2026",
}

# ---------------------------------------------------------------------------
# Member data extracted from the Excel file
# ---------------------------------------------------------------------------
MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUNE", "JULY", "AUG", "SEPT", "OCT", "NOV", "DEC"]

# Members who left on 7 July 2026 (green in original sheet) – no further penalties
LEFT_MEMBERS = {"ALPHONCE OCHIENG", "FRANKLINE OKOTI", "THOMAS ODHIAMBO"}

MEMBERS = [
    {
        "no": 1,
        "name": "JOHNSTONE MBUVI",
        "prev_total": 36000,
        "payments": {
            "JAN": None, "FEB": None, "MAR": None, "APR": None,
            "MAY": None, "JUNE": None, "JULY": None, "AUG": None,
            "SEPT": None, "OCT": None, "NOV": None, "DEC": None,
        },
        "total": 36000,
        "balance": 32200,
        "left": False,
    },
    {
        "no": 2,
        "name": "FAITH ACHIENG",
        "prev_total": 51200,
        "payments": {
            "JAN": 2000, "FEB": 2000, "MAR": None, "APR": None,
            "MAY": 2000, "JUNE": None, "JULY": None, "AUG": None,
            "SEPT": None, "OCT": None, "NOV": None, "DEC": None,
        },
        "total": 57200,
        "balance": 16100,
        "left": False,
    },
    {
        "no": 3,
        "name": "BRAYAN MBOGO",
        "prev_total": 50300,
        "payments": {
            "JAN": 2000, "FEB": 2000, "MAR": 2000, "APR": 2000,
            "MAY": 2000, "JUNE": 2000, "JULY": 2000, "AUG": 2000,
            "SEPT": 2300, "OCT": 2000, "NOV": None, "DEC": None,
        },
        "total": 70600,
        "balance": 2300,
        "left": False,
    },
    {
        "no": 4,
        "name": "DANIEL GITHANGA",
        "prev_total": 44500,
        "payments": {
            "JAN": None, "FEB": None, "MAR": None, "APR": None,
            "MAY": None, "JUNE": None, "JULY": None, "AUG": None,
            "SEPT": None, "OCT": None, "NOV": None, "DEC": None,
        },
        "total": 44500,
        "balance": 35900,
        "left": False,
    },
    {
        "no": 5,
        "name": "ALPHONCE OCHIENG",
        "prev_total": 19500,
        "payments": {
            "JAN": None, "FEB": None, "MAR": None, "APR": None,
            "MAY": None, "JUNE": None, "JULY": None, "AUG": None,
            "SEPT": None, "OCT": None, "NOV": None, "DEC": None,
        },
        "total": 19500,
        "balance": 51500,
        "left": True,  # left 7 July 2026
    },
    {
        "no": 6,
        "name": "FREDRICK MAINA",
        "prev_total": 55000,
        "payments": {
            "JAN": None, "FEB": None, "MAR": None, "APR": None,
            "MAY": 2000, "JUNE": 2000, "JULY": None, "AUG": None,
            "SEPT": None, "OCT": None, "NOV": None, "DEC": None,
        },
        "total": 59000,
        "balance": 16100,
        "left": False,
    },
    {
        "no": 7,
        "name": "MARGARET KIRIGO",
        "prev_total": 25000,
        "payments": {
            "JAN": None, "FEB": None, "MAR": None, "APR": None,
            "MAY": None, "JUNE": None, "JULY": None, "AUG": None,
            "SEPT": None, "OCT": None, "NOV": None, "DEC": None,
        },
        "total": 25000,
        "balance": 48900,
        "left": False,
    },
    {
        "no": 8,
        "name": "FRANKLINE OKOTI",
        "prev_total": 34000,
        "payments": {
            "JAN": None, "FEB": None, "MAR": None, "APR": None,
            "MAY": None, "JUNE": None, "JULY": None, "AUG": None,
            "SEPT": None, "OCT": None, "NOV": None, "DEC": None,
        },
        "total": 34000,
        "balance": 30000,
        "left": True,  # left 7 July 2026
    },
    {
        "no": 9,
        "name": "EMMANUEL BOSIRE",
        "prev_total": 16000,
        "payments": {
            "JAN": None, "FEB": None, "MAR": None, "APR": None,
            "MAY": None, "JUNE": None, "JULY": None, "AUG": None,
            "SEPT": None, "OCT": None, "NOV": None, "DEC": None,
        },
        "total": 16000,
        "balance": 49800,
        "left": False,
    },
    {
        "no": 10,
        "name": "THOMAS ODHIAMBO",
        "prev_total": 52000,
        "payments": {
            "JAN": 2000, "FEB": 2000, "MAR": None, "APR": None,
            "MAY": None, "JUNE": None, "JULY": None, "AUG": None,
            "SEPT": None, "OCT": None, "NOV": None, "DEC": None,
        },
        "total": 56000,
        "balance": 9800,
        "left": True,  # left 7 July 2026
    },
    {
        "no": 11,
        "name": "JOHN NJUGUNA",
        "prev_total": 54700,
        "payments": {
            "JAN": 2000, "FEB": 2300, "MAR": 2300, "APR": 2300,
            "MAY": 2300, "JUNE": 2300, "JULY": 2300, "AUG": 2000,
            "SEPT": None, "OCT": None, "NOV": None, "DEC": None,
        },
        "total": 72500,
        "balance": 2300,
        "left": False,
    },
    {
        "no": 12,
        "name": "STEPHEN YIEMBE",
        "prev_total": 50000,
        "payments": {
            "JAN": 2000, "FEB": 2000, "MAR": 2000, "APR": 2000,
            "MAY": 2000, "JUNE": 2000, "JULY": 2300, "AUG": 2300,
            "SEPT": 2300, "OCT": 1750, "NOV": None, "DEC": None,
        },
        "total": 70650,
        "balance": 0,
        "left": False,
    },
    {
        "no": 13,
        "name": "HENRY OGADA",
        "prev_total": 52330,
        "payments": {
            "JAN": 2000, "FEB": 2000, "MAR": None, "APR": 2000,
            "MAY": 2000, "JUNE": None, "JULY": None, "AUG": None,
            "SEPT": None, "OCT": None, "NOV": None, "DEC": None,
        },
        "total": 60330,
        "balance": 10500,
        "left": False,
    },
    {
        "no": 14,
        "name": "GILBERT CHERUIYOT",
        "prev_total": 56600,
        "payments": {
            "JAN": 2300, "FEB": 2300, "MAR": 2300, "APR": 2300,
            "MAY": 2300, "JUNE": 2300, "JULY": 2300, "AUG": 2300,
            "SEPT": 2300, "OCT": 2000, "NOV": None, "DEC": None,
        },
        "total": 79300,
        "balance": 0,
        "left": False,
    },
]

DIVIDENDS = 49530
CUMULATIVE_PREV = 646660
CUMULATIVE_TOTAL = 736310
CUMULATIVE_BALANCE = 305400

MONTHLY_TARGET = 2000
PENALTY = 300


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
    """Return status for a payment cell: paid, late (penalty), missed, n/a (left)."""
    val = member["payments"].get(month)
    if member["left"] and month in ("AUG", "SEPT", "OCT", "NOV", "DEC"):
        return "left", None
    if val is None:
        return "missed", None
    if val >= MONTHLY_TARGET + PENALTY:
        return "late", val  # paid with penalty
    if val == MONTHLY_TARGET:
        return "paid", val
    if val > 0:
        return "partial", val
    return "missed", None


def member_statement_lines(member):
    """Build human-readable statement lines for a member."""
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
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("You have been logged out.", "info")
    return redirect(url_for("login"))


@app.route("/dashboard")
@login_required
def dashboard():
    total_paid = sum(m["total"] for m in MEMBERS) + DIVIDENDS
    total_balance = sum(m["balance"] for m in MEMBERS)
    active_members = [m for m in MEMBERS if not m["left"]]
    left_members = [m for m in MEMBERS if m["left"]]
    return render_template(
        "dashboard.html",
        members=MEMBERS,
        months=MONTHS,
        dividends=DIVIDENDS,
        cumulative_prev=CUMULATIVE_PREV,
        cumulative_total=CUMULATIVE_TOTAL,
        cumulative_balance=CUMULATIVE_BALANCE,
        total_paid=total_paid,
        total_balance=total_balance,
        active_count=len(active_members),
        left_count=len(left_members),
        monthly_target=MONTHLY_TARGET,
        penalty=PENALTY,
    )


@app.route("/member/<int:no>")
@login_required
def member_detail(no):
    member = get_member(no)
    if not member:
        flash("Member not found.", "danger")
        return redirect(url_for("dashboard"))
    statement = member_statement_lines(member)
    return render_template(
        "member.html",
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
    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype="text/plain",
    )


@app.route("/download/group")
@login_required
def download_group_statement():
    """Full group statement as CSV."""
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
        buffer,
        as_attachment=True,
        download_name=f"group_contribution_statement_2026_{datetime.now().strftime('%Y%m%d')}.csv",
        mimetype="text/csv",
    )


@app.route("/download/group-txt")
@login_required
def download_group_txt():
    """Full group statement as plain text."""
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
        buffer,
        as_attachment=True,
        download_name=f"group_full_statement_2026_{datetime.now().strftime('%Y%m%d')}.txt",
        mimetype="text/plain",
    )


if __name__ == "__main__":
    # For local development
    app.run(host="0.0.0.0", port=5000, debug=True)
