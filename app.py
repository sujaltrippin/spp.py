from flask import Flask, request, jsonify
from automation import process_single_expense

app = Flask(__name__)

@app.route('/log-expense', methods=['POST'])
def log_expense():
    try:
        data = request.json

        booking_id = data.get("booking_id")
        vendor_name = data.get("vendor_name")
        property_name = data.get("property_name")
        amount = data.get("amount")
        sub_category = data.get("sub")

        if not all([booking_id, vendor_name, property_name, amount]):
            return jsonify({"status": "error", "message": "Missing required fields"}), 400

        success = process_single_expense(
            booking_id, vendor_name, property_name, amount, sub_category
        )

        if success:
            return jsonify({"status": "success", "message": "Expense logged"}), 200
        else:
            return jsonify({"status": "error", "message": "Logging failed"}), 500

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/', methods=['GET'])
def home():
    return "StayVista Automation Flask Server Running"

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000)
