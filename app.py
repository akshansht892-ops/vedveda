"""
Vedveda Sales & Analytics Dashboard
====================================
Flask application entry point with automatic/manual Excel synchronization 
and member management.
"""

import os
from datetime import datetime
from flask import Flask, jsonify, render_template, request, abort
from flask_sqlalchemy import SQLAlchemy

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DATA_DIR  = os.environ.get("VEDVEDA_DATA_DIR", BASE_DIR)
RES_DIR   = os.environ.get("VEDVEDA_RES_DIR",  BASE_DIR)

app = Flask(
    __name__,
    template_folder = os.path.join(RES_DIR, "templates"),
    static_folder   = os.path.join(RES_DIR, "static"),
)
app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{os.path.join(DATA_DIR, 'vedveda.db')}"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["JSON_SORT_KEYS"] = False

db = SQLAlchemy(app)

_LAST_EXCEL_MTIME = 0

def check_and_sync_excel():
    """Check if VEDVEDA.xlsx has been modified and re-seed if necessary."""
    global _LAST_EXCEL_MTIME
    xlsx_path = os.path.join(RES_DIR, "VEDVEDA.xlsx")
    if not os.path.exists(xlsx_path):
        xlsx_path = os.path.join(BASE_DIR, "VEDVEDA.xlsx")
    
    if os.path.exists(xlsx_path):
        mtime = os.path.getmtime(xlsx_path)
        if _LAST_EXCEL_MTIME == 0:
            _LAST_EXCEL_MTIME = mtime
        elif mtime > _LAST_EXCEL_MTIME:
            print("[INFO] VEDVEDA.xlsx change detected. Re-syncing database...")
            try:
                from seed import seed_database
                db.drop_all()
                db.create_all()
                seed_database(db, Member, Transaction)
                _LAST_EXCEL_MTIME = mtime
                print("[INFO] Excel sync complete.")
            except Exception as e:
                print(f"[ERROR] Excel sync failed: {e}")

class Member(db.Model):
    __tablename__ = "members"

    id   = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False, unique=True)
    role = db.Column(db.String(100), nullable=False, default="Team Member")

    transactions = db.relationship("Transaction", backref="member",
                                   lazy=True, cascade="all, delete-orphan")

    def to_dict(self):
        txns = self.transactions
        total_pv     = sum(t.pv or 0 for t in txns)
        total_amount = sum(t.amount or 0 for t in txns)
        paid_amount  = sum(t.amount or 0 for t in txns if t.payment_status == "Paid")
        pending_amt  = sum(t.amount or 0 for t in txns if t.payment_status == "Pending")
        penal_pv     = sum(t.penal_pv or 0 for t in txns)
        penal_amount = sum(t.penal_amount or 0 for t in txns)
        return {
            "id":            self.id,
            "name":          self.name,
            "role":          self.role,
            "total_pv":      round(total_pv, 2),
            "total_amount":  round(total_amount, 2),
            "paid_amount":   round(paid_amount, 2),
            "pending_amount": round(pending_amt, 2),
            "penal_pv":      round(penal_pv, 2),
            "penal_amount":  round(penal_amount, 2),
            "transaction_count": len(txns),
        }


class Transaction(db.Model):
    __tablename__ = "transactions"

    id             = db.Column(db.Integer, primary_key=True)
    member_id      = db.Column(db.Integer, db.ForeignKey("members.id"), nullable=False)
    date           = db.Column(db.Date, nullable=True)
    product_name   = db.Column(db.String(200), nullable=False)
    quantity       = db.Column(db.Integer, nullable=True, default=1)
    pv             = db.Column(db.Float, nullable=True, default=0)
    amount         = db.Column(db.Float, nullable=True, default=0)
    payment_status = db.Column(db.String(20), nullable=False, default="Pending")
    penal_pv       = db.Column(db.Float, nullable=True, default=0)
    penal_amount   = db.Column(db.Float, nullable=True, default=0)

    def to_dict(self):
        return {
            "id":             self.id,
            "member_id":      self.member_id,
            "member_name":    self.member.name if self.member else None,
            "date":           self.date.isoformat() if self.date else None,
            "product_name":   self.product_name,
            "quantity":       self.quantity,
            "pv":             self.pv,
            "amount":         self.amount,
            "payment_status": self.payment_status,
            "penal_pv":       self.penal_pv,
            "penal_amount":   self.penal_amount,
        }

def _parse_date(val):
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.date()
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(str(val), fmt).date()
        except ValueError:
            pass
    return None

@app.before_request
def before_request_sync():
    check_and_sync_excel()

@app.route("/")
def index():
    check_and_sync_excel()
    txns = Transaction.query.all()
    global_pv      = round(sum(t.pv or 0 for t in txns), 2)
    global_revenue = round(sum(t.amount or 0 for t in txns if t.payment_status == "Paid"), 2)
    global_pending = round(sum(t.amount or 0 for t in txns if t.payment_status == "Pending"), 2)
    members        = Member.query.order_by(Member.name).all()

    return render_template(
        "dashboard.html",
        global_pv=global_pv,
        global_revenue=global_revenue,
        global_pending=global_pending,
        members=[m.to_dict() for m in members],
    )

@app.route("/api/sync-excel", methods=["POST"])
def api_sync_excel():
    try:
        from seed import seed_database
        db.drop_all()
        db.create_all()
        seed_database(db, Member, Transaction)
        return jsonify({"success": True, "message": "Successfully synced with VEDVEDA.xlsx!"})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/api/members", methods=["GET"])
def api_members():
    check_and_sync_excel()
    members = Member.query.order_by(Member.name).all()
    return jsonify([m.to_dict() for m in members])

@app.route("/api/members", methods=["POST"])
def api_add_member():
    data = request.get_json(force=True)
    if not data or not data.get("name"):
        abort(400, description="'name' is required.")
    if Member.query.filter_by(name=data["name"].strip()).first():
        abort(409, description="A member with that name already exists.")
    member = Member(
        name=data["name"].strip(),
        role=data.get("role", "Team Member").strip(),
    )
    db.session.add(member)
    db.session.commit()
    return jsonify(member.to_dict()), 201

@app.route("/api/member/<int:member_id>", methods=["GET"])
def api_member_detail(member_id):
    check_and_sync_excel()
    member = db.session.get(Member, member_id)
    if not member:
        abort(404, description="Member not found.")

    txns = (Transaction.query
            .filter_by(member_id=member_id)
            .order_by(Transaction.date.desc())
            .all())

    product_breakdown = {}
    for t in txns:
        p = t.product_name
        if p not in product_breakdown:
            product_breakdown[p] = {"pv": 0, "amount": 0, "count": 0}
        product_breakdown[p]["pv"]     += t.pv or 0
        product_breakdown[p]["amount"] += t.amount or 0
        product_breakdown[p]["count"]  += 1

    return jsonify({
        "member":            member.to_dict(),
        "transactions":      [t.to_dict() for t in txns],
        "product_breakdown": product_breakdown,
    })

@app.route("/api/transactions", methods=["GET"])
def api_transactions():
    check_and_sync_excel()
    member_id = request.args.get("member_id", type=int)
    query = Transaction.query.order_by(Transaction.date.desc())
    if member_id:
        query = query.filter_by(member_id=member_id)
    return jsonify([t.to_dict() for t in query.all()])

@app.route("/api/transactions", methods=["POST"])
def api_add_transaction():
    data = request.get_json(force=True)
    required = ("member_id", "product_name")
    for field in required:
        if not data.get(field):
            abort(400, description=f"'{field}' is required.")

    member = db.session.get(Member, data["member_id"])
    if not member:
        abort(404, description="Member not found.")

    status = data.get("payment_status", "Pending")
    if status not in ("Paid", "Pending"):
        abort(400, description="payment_status must be 'Paid' or 'Pending'.")

    txn = Transaction(
        member_id      = data["member_id"],
        date           = _parse_date(data.get("date")),
        product_name   = data["product_name"].strip(),
        quantity       = data.get("quantity", 1),
        pv             = float(data.get("pv") or 0),
        amount         = float(data.get("amount") or 0),
        payment_status = status,
        penal_pv       = float(data.get("penal_pv") or 0),
        penal_amount   = float(data.get("penal_amount") or 0),
    )
    db.session.add(txn)
    db.session.commit()
    return jsonify(txn.to_dict()), 201

@app.route("/api/transactions/<int:txn_id>", methods=["PUT"])
def api_update_transaction(txn_id):
    txn = db.session.get(Transaction, txn_id)
    if not txn:
        abort(404, description="Transaction not found.")

    data = request.get_json(force=True)

    if "date" in data:
        txn.date = _parse_date(data["date"])
    if "product_name" in data and data["product_name"]:
        txn.product_name = data["product_name"].strip()
    if "quantity" in data:
        txn.quantity = data["quantity"]
    if "pv" in data:
        txn.pv = float(data["pv"] or 0)
    if "amount" in data:
        txn.amount = float(data["amount"] or 0)
    if "payment_status" in data:
        status = data["payment_status"]
        if status not in ("Paid", "Pending"):
            abort(400, description="payment_status must be 'Paid' or 'Pending'.")
        txn.payment_status = status
    if "penal_pv" in data:
        txn.penal_pv = float(data["penal_pv"] or 0)
    if "penal_amount" in data:
        txn.penal_amount = float(data["penal_amount"] or 0)
    if "member_id" in data:
        member = db.session.get(Member, data["member_id"])
        if not member:
            abort(404, description="Member not found.")
        txn.member_id = data["member_id"]

    db.session.commit()
    return jsonify(txn.to_dict())

@app.route("/api/transactions/<int:txn_id>", methods=["DELETE"])
def api_delete_transaction(txn_id):
    txn = db.session.get(Transaction, txn_id)
    if not txn:
        abort(404, description="Transaction not found.")
    db.session.delete(txn)
    db.session.commit()
    return jsonify({"success": True, "deleted_id": txn_id})

@app.route("/api/stats/overview", methods=["GET"])
def api_stats_overview():
    check_and_sync_excel()
    members = Member.query.order_by(Member.name).all()

    pv_per_member = [
        {"member": m.name, "pv": round(sum(t.pv or 0 for t in m.transactions), 2)}
        for m in members
    ]

    txns = Transaction.query.all()
    paid_total    = round(sum(t.amount or 0 for t in txns if t.payment_status == "Paid"), 2)
    pending_total = round(sum(t.amount or 0 for t in txns if t.payment_status == "Pending"), 2)

    monthly = {}
    for t in txns:
        if t.date:
            key = t.date.strftime("%Y-%m")
            monthly[key] = monthly.get(key, 0) + (t.amount or 0)
    monthly_sorted = dict(sorted(monthly.items()))

    return jsonify({
        "pv_per_member": pv_per_member,
        "paid_vs_pending": {
            "paid":    paid_total,
            "pending": pending_total,
        },
        "monthly_revenue": monthly_sorted,
    })

if __name__ == "__main__":
    with app.app_context():
        db.create_all()
        check_and_sync_excel()
    app.run(debug=True, host="0.0.0.0", port=5000)