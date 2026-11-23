import urllib.request
import json
import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(__file__), "focus_portal.db")

def test_push_lead():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT value FROM settings WHERE key = 'webhook_api_key'")
    row = cursor.fetchone()
    if not row:
        print("API Key not found in DB! Please start app.py first.")
        return
    api_key = row[0]
    conn.close()

    print(f"Using Webhook API Key: {api_key}")

    url = "http://127.0.0.1:5500/api/webhook/lead"
    payload = {
        "name": "Jagjit Singh",
        "phone": "9814012345",
        "email": "jagjit@example.com",
        "course": "PTE Academic - Morning Batch",
        "city": "Amritsar",
        "source": "Main Admission Portal Online Form"
    }

    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "X-API-Key": api_key
        },
        method="POST"
    )

    try:
        with urllib.request.urlopen(req) as response:
            res_data = json.loads(response.read().decode("utf-8"))
            print("Webhook Test Response:")
            print(json.dumps(res_data, indent=2))
    except Exception as e:
        print(f"Error testing webhook: {e}")

if __name__ == "__main__":
    test_push_lead()
