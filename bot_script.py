import time
import os
import firebase_admin
from firebase_admin import credentials, firestore
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==========================================
# ส่วนที่ 1: จัดการ Firestore (ระบบ Lock)
# ==========================================
def initialize_firestore():
    try:
        if not os.path.exists("serviceAccount.json"):
            print("❌ ไม่พบไฟล์ serviceAccount.json")
            return None
        cred = credentials.Certificate("serviceAccount.json")
        firebase_admin.initialize_app(cred)
        return firestore.client()
    except Exception as e:
        print(f"❌ Error Firestore Init: {e}")
        return None

def get_and_lock_ticket(db):
    try:
        docs = db.collection("incidents") \
                 .where("status", "==", "Pending") \
                 .where("is_pulled", "==", False) \
                 .limit(1).get()
        for doc in docs:
            doc_id = doc.id
            doc_data = doc.to_dict()
            db.collection("incidents").document(doc_id).update({"is_pulled": True})
            print(f"\n🔒 ล็อกงาน {doc_id} เรียบร้อย...")
            return doc_id, doc_data
        return None, None
    except Exception as e:
        print(f"❌ Error Firestore: {e}")
        return None, None

def mark_as_completed(db, doc_id):
    try:
        db.collection("incidents").document(doc_id).update({"status": "Completed"})
        print(f"✅ [Firestore] ปรับสถานะเป็น Completed")
    except Exception as e:
        print(f"❌ Error Firestore Update: {e}")

# ==========================================
# ส่วนที่ 2: ควบคุม ServiceNow (สั้น ชับ ตาม ID จริง)
# ==========================================
def fill_servicenow_ticket(driver, wait, data):
    try:
        print(f"🚀 เริ่มสร้าง Ticket: {data.get('short_description', 'No Subject')}")
        driver.get("https://keris.service-now.com/incident.do?sys_id=-1")
        
        # ⏳ รอหน้าฟอร์มโหลด (ให้ช่อง Short Description ขึ้นมาเป็นสัญญาณ)
        wait.until(EC.presence_of_element_located((By.ID, "incident.short_description")))

        # --- 1. เลือก Demand (ใช้ ID จริงที่เพิ่งเจอ: incident.u_demand) ---
        print("- กำลังเลือก Sense & Respond Demand...")
        Select(driver.find_element(By.ID, "incident.u_demand")).select_by_visible_text("Retail Solution & Delivery")

        # --- 2. กรอก Caller (ใช้ sys_display) ---
        print("- กำลังกรอก Caller...")
        caller_field = driver.find_element(By.ID, "sys_display.incident.caller_id")
        caller_field.clear()
        caller_field.send_keys(data.get("caller", ""), Keys.RETURN)
        time.sleep(2)
        
        # --- 3. ฟิลด์มาตรฐาน (สั้นๆ แบบ Urgency) ---
        driver.find_element(By.ID, "incident.short_description").send_keys(data.get("short_description", ""))
        driver.find_element(By.ID, "incident.description").send_keys(data.get("description", ""))
        
        Select(driver.find_element(By.ID, "incident.category")).select_by_visible_text(data.get("category", "Software"))
        Select(driver.find_element(By.ID, "incident.impact")).select_by_visible_text(data.get("impact", "5 - Minor"))
        Select(driver.find_element(By.ID, "incident.urgency")).select_by_visible_text(data.get("urgency", "5 - Minor"))
        
        # --- 4. Assignment Group & Reported By ---
        driver.find_element(By.ID, "sys_display.incident.assignment_group").send_keys("FTH ISS3000 MALL", Keys.RETURN)
        time.sleep(2)

        driver.find_element(By.ID, "sys_display.incident.u_reported_by").send_keys(data.get("reported_by", "DON-001"), Keys.RETURN)
        time.sleep(2)

        # Mitigate SLA
        driver.find_element(By.ID, "incident.u_mitigate_sla_description").send_keys(data.get("contact_info", ""))

        # --- 5. กด Save ---
        print("💾 กำลังกดบันทึก...")
        driver.find_element(By.ID, "sysverb_insert").click()
        
        time.sleep(5) 
        return True 

    except Exception as e:
        print(f"❌ Error ServiceNow: {e}")
        return False

# ==========================================
# Main Loop
# ==========================================
if __name__ == "__main__":
    db = initialize_firestore()
    if db:
        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        try:
            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 15)
            print("🤖 บอทพร้อมทำงาน! กำลังเฝ้า Firestore...")

            while True:
                doc_id, doc_data = get_and_lock_ticket(db)
                if doc_id:
                    if fill_servicenow_ticket(driver, wait, doc_data):
                        mark_as_completed(db, doc_id)
                    else:
                        print(f"⚠️ งาน {doc_id} มีปัญหา (ตรวจสอบ is_pulled ใน Firestore)")
                else:
                    print(".", end="", flush=True)
                time.sleep(15)
        except Exception as e:
            print(f"❌ ไม่พบ Chrome โหมด Debug: {e}")
            input("กด Enter เพื่อปิด...")
