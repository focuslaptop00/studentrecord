# Focus Institute - Secondary Calling & Telecalling CRM Portal

Yeh Focus Institute da **Secondary Calling / CRM Portal** hai, jo khas tor te **3-5 Calling Staff Members** aur **Manager (Arsh)** layi design kita gaya hai. 

---

## 🚀 Key Features

### 1. Manager Panel (Arsh layi)
* **Excel / CSV Bulk Import:** Main portal ton direct `.xlsx` ya `.csv` file upload karke sekhan vich sekhan leads add karo.
* **Auto Round-Robin Distribution:** Naye leads automatically 3-5 staff vich barabar divide ho jande ne.
* **Staff Performance Monitoring:** Live track karo ki Gurpreet, Simran ya Aman ne kinniyan calls kitian te kinne admissions convert kite.
* **Option B Webhook API:** Main admission portal developer layi ready-made API endpoint & Secret Key.
* **Export Reports:** Sabhi leads aur remarks nu Excel / CSV format vich download karo.

### 2. Staff / Caller Panel (Calling Team layi)
* **Full Phone Number Visibility:** Phone number poora dikhega calling waste.
* **1-Click Call & WhatsApp:** Direct call dial aur pre-written greeting message naal WhatsApp chat.
* **Call Status Selection:** (*Interested*, *Call Back / Follow-up*, *Admission Done*, *Ringing*, *Switched Off*, *Not Interested*).
* **Call History Timeline:** Student di saari purani call history te notes timestamp naal check karo.

---

## 🔑 Default Login Credentials

| Role | Name | Username | Password |
|---|---|---|---|
| **Manager** | Arsh (Manager) | `arsh` | `arsh123` |
| **Staff 1** | Gurpreet Singh | `gurpreet` | `staff123` |
| **Staff 2** | Simran Kaur | `simran` | `staff123` |
| **Staff 3** | Aman Sharma | `aman` | `staff123` |

---

## 💻 How to Run

```bash
cd /home/harsh/.gemini/antigravity/scratch/focus-telecalling-portal
python3 app.py
```
Browser vich open karo:
👉 **`http://127.0.0.1:5500`**
