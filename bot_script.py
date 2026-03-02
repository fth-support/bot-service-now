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
# ส่วนที่ 1: จัดการ Firestore (ดึงจาก status: on process)
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
    print("\n🔍 กำลังตรวจสอบงานใหม่จาก Firestore (status: on process)...")
    try:
        # ปรับการค้นหาตามภาพล่าสุดของคุณ: status == "on process"
        docs = db.collection("incidents") \
                 .where("status", "==", "on process") \
                 .where("is_pulled", "==", False) \
                 .limit(1).get()
        
        for doc in docs:
            doc_id = doc.id
            doc_data = doc.to_dict()
            db.collection("incidents").document(doc_id).update({"is_pulled": True})
            print(f"🔒 ล็อกงาน {doc_id} เรียบร้อย (กำลังเริ่มสร้างใน UAT)")
            return doc_id, doc_data
        return None, None
    except Exception as e:
        print(f"❌ Error Firestore: {e}")
        return None, None

def mark_as_completed(db, doc_id):
    try:
        db.collection("incidents").document(doc_id).update({"status": "Completed"})
        print(f"✅ [Firestore] ปิดงาน {doc_id} สมบูรณ์")
    except Exception as e:
        print(f"❌ Error Firestore Update: {e}")

# ==========================================
# ส่วนที่ 2: ควบคุม ServiceNow (UAT Version)
# ==========================================
def fill_servicenow_ticket(driver, wait, data):
    try:
        print(f"🚀 เริ่มสร้าง Ticket ในระบบ UAT: {data.get('description', 'No Subject')}")
        # เปลี่ยนเป็น URL UAT ตามที่คุณระบุ (ใช้ Raw Form เพื่อความเสถียร)
        uat_form_url = "https://keristest.service-now.com/incident.do?sys_id=-1"
        driver.get(uat_form_url)
        
        # รอให้หน้าฟอร์มกางตัวออก
        time.sleep(6) 

        # --- จุดที่ 1: กรอกแต่ Short Description (ดึงจาก description ใน Firebase) ---
        print("- กำลังกรอก Short Description...")
        driver.find_element(By.ID, "incident.short_description").clear()
        driver.find_element(By.ID, "incident.short_description").send_keys(data.get("description", ""))

        # --- ฟิลด์บังคับอื่นๆ (Caller) ---
        print("- กำลังกรอก Caller...")
        caller_field = driver.find_element(By.ID, "sys_display.incident.caller_id")
        caller_field.clear()
        caller_field.send_keys(data.get("caller", ""), Keys.RETURN)
        time.sleep(2)

        # --- จุดที่ 2: กรอก External Ref No 4 (ticket_id จาก Firebase) ---
        print("- กำลังค้นหาและกรอก External Ref No 4...")
        try:
            # ปกติ ServiceNow ฟิลด์ใน Tab จะโหลดมาใน DOM อยู่แล้ว 
            # ID มักจะขึ้นต้นด้วย incident.u_... รบกวนคุณ Inspect ดู ID จริงอีกครั้งถ้าตัวนี้ไม่ทำงาน
            # ผมใช้ Selector แบบกวาดหาคำว่า external_ref_no_4
            ext_ref_field = driver.find_element(By.CSS_SELECTOR, "input[id*='external_ref_no_4']")
            ext_ref_field.clear()
            ext_ref_field.send_keys(data.get("ticket_id", ""))
            print("  ✅ กรอก External Ref No 4 สำเร็จ")
        except:
            print("  ⚠️ คำเตือน: หาช่อง External Ref No 4 ไม่เจอ (อาจต้องคลิก Tab ก่อน หรือ ID ไม่ตรง)")

        # --- 5. กด Save ---
        print("💾 กำลังกดบันทึก...")
        driver.find_element(By.ID, "sysverb_insert").click()
        
        time.sleep(5) 
        return True 

    except Exception as e:
        print(f"❌ Error ServiceNow UAT: {e}")
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
            print("🤖 บอทพร้อมทำงานบนระบบ UAT! (เฝ้าดูสถานะ on process)")

            while True:
                doc_id, doc_data = get_and_lock_ticket(db)
                if doc_id:
                    if fill_servicenow_ticket(driver, wait, doc_data):
                        mark_as_completed(db, doc_id)
                    else:
                        print(f"⚠️ งาน {doc_id} มีปัญหาบนเว็บ UAT")
                else:
                    print(".", end="", flush=True)
                time.sleep(15)
        except Exception as e:
            print(f"❌ ไม่พบ Chrome โหมด Debug (Port 9222): {e}")
            input("กด Enter เพื่อปิด...")
