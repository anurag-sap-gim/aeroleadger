import datetime
import os
import sys
import tempfile
import threading
import time
from decimal import Decimal

from flask import Flask, jsonify, render_template, request, send_file, abort

sys.path.insert(0, os.path.dirname(__file__))
import airline_accounting as aa

app = Flask(__name__)


@app.context_processor
def inject_now():
    return {"now": datetime.datetime.now()}

# ---------------------------------------------------------------------------
# Shared state — updated by background FX thread
# ---------------------------------------------------------------------------
_lock   = threading.Lock()
_fx     = aa.FXRateManager(use_live=True)
_ledger = aa.LedgerBook()
_leases = [
    aa.IFRS16LeaseSchedule(aa.LeaseConfig(
        lease_id="LEASE-A320-01", aircraft_type="Airbus A320neo",
        commencement=datetime.date(datetime.date.today().year - 2, 4, 1),
        lease_term_months=120, monthly_payment=Decimal("210000"),
        effective_rate=Decimal("0.065"))),
    aa.IFRS16LeaseSchedule(aa.LeaseConfig(
        lease_id="LEASE-B737-01", aircraft_type="Boeing 737 MAX 8",
        commencement=datetime.date(datetime.date.today().year - 1, 7, 1),
        lease_term_months=96, monthly_payment=Decimal("185000"),
        effective_rate=Decimal("0.060"))),
]
_state = {"rates": dict(aa._BASE_RATES), "source": "SIM",
          "updated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}


def _seed_ledger():
    rates = aa._BASE_RATES
    today = datetime.date.today()
    base  = datetime.date(today.year, 1, 15)
    txs = [
        ("Jet fuel purchase - JFK", "Fuel Expense", "EXPENSE", 480000, 0, "USD", rates["USD"], True, 0),
        ("Fuel AP accrual", "Fuel Payable", "LIABILITY", 0, 480000, "USD", rates["USD"], True, 0),
        ("Lease payment A320neo Q1", "Lease Liability", "LIABILITY", 210000, 0, "EUR", Decimal("1"), True, 2),
        ("Lease interest A320neo", "Interest Expense", "EXPENSE", 14300, 0, "EUR", Decimal("1"), False, 2),
        ("ROU depreciation A320neo", "ROU Asset Depreciation", "EXPENSE", 38500, 0, "EUR", Decimal("1"), False, 2),
        ("Heavy maintenance reserve", "Maintenance Reserve", "EXPENSE", 95000, 0, "GBP", rates["GBP"], False, 3),
        ("Maintenance payable", "Maintenance Payable", "LIABILITY", 0, 95000, "GBP", rates["GBP"], True, 3),
        ("Passenger revenue LHR-SIN", "Passenger Revenue", "REVENUE", 0, 620000, "SGD", rates["SGD"], False, 4),
        ("Cargo revenue BOM-FRA", "Cargo Revenue", "REVENUE", 0, 8500000, "INR", rates["INR"], False, 4),
        ("Passenger revenue MUC-DXB", "Passenger Revenue", "REVENUE", 0, 380000, "USD", rates["USD"], False, 5),
        ("Staff salaries flight crew", "Staff Costs", "EXPENSE", 340000, 0, "EUR", Decimal("1"), False, 6),
        ("Airport charges CDG", "Airport Charges", "EXPENSE", 12400, 0, "EUR", Decimal("1"), False, 7),
        ("ATC charges Eurocontrol", "ATC Charges Payable", "LIABILITY", 0, 7200, "EUR", Decimal("1"), True, 7),
        ("Ground handling Changi", "Ground Handling", "EXPENSE", 42000, 0, "SGD", rates["SGD"], False, 8),
        ("Insurance premium fleet", "Insurance Expense", "EXPENSE", 56000, 0, "EUR", Decimal("1"), False, 9),
        ("Catering costs economy", "Catering Costs", "EXPENSE", 28400, 0, "EUR", Decimal("1"), False, 9),
        ("Fuel hedge gain", "Fuel Hedge Gain", "REVENUE", 0, 22000, "USD", rates["USD"], False, 10),
        ("Navigation fees Eurocontrol", "Navigation Fees", "EXPENSE", 9800, 0, "EUR", Decimal("1"), False, 11),
        ("B737 lease payment Q1", "Lease Liability", "LIABILITY", 185000, 0, "EUR", Decimal("1"), True, 12),
        ("B737 lease interest Q1", "Interest Expense", "EXPENSE", 10800, 0, "EUR", Decimal("1"), False, 12),
    ]
    for desc, acct, atype, dr, cr, ccy, rate, mon, day_off in txs:
        d = base + datetime.timedelta(days=day_off)
        _ledger.add_transaction(aa.Transaction(
            tx_id=_ledger.next_tx_id(), date=d, description=desc,
            account=acct, account_type=atype,
            debit_eur=Decimal(str(dr)), credit_eur=Decimal(str(cr)),
            currency=ccy, fx_rate=Decimal(str(rate)), is_monetary=mon))


_seed_ledger()


def _fx_refresh_loop():
    while True:
        try:
            rates = _fx.fetch_or_simulate()
            _fx.record_rates(rates)
            with _lock:
                _state["rates"]   = {k: float(v) for k, v in rates.items()}
                _state["source"]  = "Live API (frankfurter.dev)" if _fx.is_live() else "Simulated"
                _state["updated"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            pass
        time.sleep(15)


_thread = threading.Thread(target=_fx_refresh_loop, daemon=True)
_thread.start()

# ---------------------------------------------------------------------------
# Pages
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/calculator")
def calculator():
    return render_template("calculator.html",
                           today=datetime.date.today().isoformat())


@app.route("/about")
def about():
    return render_template("about.html")

# ---------------------------------------------------------------------------
# API
# ---------------------------------------------------------------------------

@app.route("/api/fx-rates")
def api_fx_rates():
    with _lock:
        data = dict(_state)
    return jsonify({
        "rates":   data["rates"],
        "source":  data["source"],
        "updated": data["updated"],
        "base":    "EUR",
    })


@app.route("/api/summary")
def api_summary():
    ta  = _ledger.total_assets()
    tl  = _ledger.total_liabilities()
    rev = _ledger.total_revenue()
    exp = _ledger.total_expenses()
    ni  = _ledger.net_income()
    ll  = sum(ls.current_liability() for ls in _leases)
    rou = sum(ls.current_rou()       for ls in _leases)
    return jsonify({
        "total_assets":       float(ta),
        "total_liabilities":  float(tl),
        "net_equity":         float(ta - tl),
        "total_revenue":      float(rev),
        "total_expenses":     float(exp),
        "net_income":         float(ni),
        "lease_liability":    float(ll),
        "rou_asset":          float(rou),
        "is_profit":          ni >= 0,
    })


@app.route("/api/lease-calc", methods=["POST"])
def api_lease_calc():
    body = request.get_json(force=True) or {}
    try:
        monthly_payment   = Decimal(str(body["monthly_payment"]))
        lease_term_months = int(body["lease_term_months"])
        effective_rate    = Decimal(str(body["effective_rate"]))
        aircraft_type     = str(body.get("aircraft_type", "Aircraft"))
        comm_str          = body.get("commencement", datetime.date.today().isoformat())
        commencement      = datetime.date.fromisoformat(comm_str)
    except (KeyError, ValueError) as e:
        return jsonify({"error": str(e)}), 400

    if not (1 <= lease_term_months <= 360):
        return jsonify({"error": "lease_term_months must be 1–360"}), 400
    if not (Decimal("0.001") <= effective_rate <= Decimal("0.5")):
        return jsonify({"error": "effective_rate must be 0.001–0.5 (e.g. 0.085 for 8.5%)"}), 400
    if monthly_payment <= 0:
        return jsonify({"error": "monthly_payment must be positive"}), 400

    cfg = aa.LeaseConfig(
        lease_id="CALC", aircraft_type=aircraft_type,
        commencement=commencement,
        lease_term_months=lease_term_months,
        monthly_payment=monthly_payment,
        effective_rate=effective_rate,
    )
    ls = aa.IFRS16LeaseSchedule(cfg)
    sched = ls.compute_schedule()

    initial_pv       = ls.initial_pv()
    monthly_rate     = ls.monthly_rate()
    current_liab     = ls.current_liability()
    current_rou      = ls.current_rou()
    ytd_interest     = ls.cumulative_interest_ytd()
    total_interest   = sum(r["interest_charge"] for r in sched)
    total_payments   = monthly_payment * lease_term_months
    total_cost       = total_payments
    roi_pct          = float((initial_pv - total_interest) / initial_pv * 100) if initial_pv != 0 else 0.0
    is_profit        = total_interest < initial_pv * Decimal("0.3")

    first3 = [
        {"period": r["period"],
         "date":   r["date"].strftime("%d-%b-%Y"),
         "opening_liability": float(r["opening_liability"]),
         "interest":  float(r["interest_charge"]),
         "principal": float(r["principal"]),
         "closing_liability": float(r["closing_liability"]),
         "rou_closing": float(r["rou_closing"])}
        for r in sched[:3]
    ]
    last3 = [
        {"period": r["period"],
         "date":   r["date"].strftime("%d-%b-%Y"),
         "opening_liability": float(r["opening_liability"]),
         "interest":  float(r["interest_charge"]),
         "principal": float(r["principal"]),
         "closing_liability": float(r["closing_liability"]),
         "rou_closing": float(r["rou_closing"])}
        for r in sched[-3:]
    ]

    return jsonify({
        "aircraft_type":        aircraft_type,
        "lease_term_months":    lease_term_months,
        "effective_rate_pct":   float(effective_rate * 100),
        "initial_pv":           float(initial_pv),
        "monthly_rate_pct":     float(monthly_rate * 100),
        "total_payments":       float(total_payments),
        "total_interest":       float(total_interest),
        "interest_ratio_pct":   float(total_interest / initial_pv * 100) if initial_pv else 0,
        "current_liability":    float(current_liab),
        "current_rou":          float(current_rou),
        "ytd_interest":         float(ytd_interest),
        "roi_pct":              roi_pct,
        "verdict":              "PROFITABLE" if is_profit else "HIGH COST",
        "verdict_detail": (
            "Interest cost is within acceptable range (<30% of asset value). "
            "Lease is financially sound at this IBR."
        ) if is_profit else (
            "Interest cost exceeds 30% of initial asset value. "
            "Consider renegotiating IBR or shorter term."
        ),
        "first_3_periods": first3,
        "last_3_periods":  last3,
    })


@app.route("/download/excel")
def download_excel():
    try:
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp.close()
        writer = aa.WorkbookWriter(_ledger, _fx, _leases)
        writer.write(tmp.name)
        filename = f"AeroLedger_{datetime.date.today().isoformat()}.xlsx"
        return send_file(tmp.name, as_attachment=True,
                         download_name=filename,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        abort(500, description=str(e))


if __name__ == "__main__":
    app.run(debug=True, port=5000)
