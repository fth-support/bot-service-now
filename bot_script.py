import time
import os
import firebase_admin
from firebase_admin import credentials, firestore
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==========================================
# ส่วนที่ 1: จัดการ Firestore (System Lock)
# ==========================================
def initialize_firestore():
    try:
        if not os.path.exists("serviceAccount.json"):
            print("❌ ไม่พบไฟล์ serviceAccount.json ในโฟลเดอร์เดียวกับบอท")
            return None
        cred = credentials.Certificate("serviceAccount.json")
        firebase_admin.initialize_app(cred)
        return firestore.client()
    except Exception as e:
        print(f"❌ Error Firestore Init: {e}")
        return None

def get_and_lock_ticket(db):
    print("\n🔍 ตรวจสอบงานใหม่ (on process)...")
    try:
        # ค้นหางานที่ status == 'on process' และยังไม่ถูกดึงไป
        docs = db.collection("incidents") \
                 .where("status", "==", "on process") \
                 .where("is_pulled", "==", False) \
                 .limit(1).get()
        
        for doc in docs:
            doc_id = doc.id
            doc_data = doc.to_dict()
            # ทำการล็อกงานทันทีเพื่อกันรันซ้ำ
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
        print(f"✅ [Firestore] อัปเดตสถานะเป็น Completed")
    except Exception as e:
        print(f"❌ Error Firestore Update: {e}")

# ==========================================
# ส่วนที่ 2: ควบคุม ServiceNow UAT
# ==========================================
def fill_servicenow_ticket(driver, wait, data):
    try:
        print(f"🚀 เริ่มสร้าง Ticket: {data.get('description', 'No Subject')}")
        
        # บังคับบอทให้มาทำงานที่ Tab ล่าสุดที่เปิดอยู่
        all_handles = driver.window_handles
        driver.switch_to.window(all_handles[-1])
        
        # เข้าหน้า Raw Form ของ UAT โดยตรง
        uat_form_url = "https://keristest.service-now.com/incident.do?sys_id=-1"
        driver.get(uat_form_url)
        time.sleep(5) 

        # 1. กรอก Short Description (ใช้ค่าจาก description ใน Firebase)
        print("- กรอก Short Description...")
        short_desc = wait.until(EC.presence_of_element_located((By.ID, "incident.short_description")))
        short_desc.clear()
        short_desc.send_keys(data.get("description", ""))

        # 2. กรอก Caller
        print("- กรอก Caller...")
        caller_field = driver.find_element(By.ID, "sys_display.incident.caller_id")
        caller_field.clear()
        caller_field.send_keys(data.get("caller", ""), Keys.RETURN)
        time.sleep(2)

        # 3. กรอก External Ref No 4 (ใช้ค่าจาก ticket_id ใน Firebase)
        print("- กรอก External Ref No 4...")
        try:
            # ใช้ CSS Selector แบบกวาดหา ID ที่เกี่ยวข้องกับ external_ref_no_4
            ext_ref = driver.find_element(By.CSS_SELECTOR, "input[id*='external_ref_no_4']")
            ext_ref.clear()
            ext_ref.send_keys(data.get("ticket_id", ""))
            print("  ✅ กรอก External Ref No 4 สำเร็จ")
        except:
            print("  ⚠️ หาช่อง External Ref No 4 ไม่เจอ")

        # 4. กด Save (Submit)
        print("💾 กำลังกดบันทึก...")
        driver.find_element(By.ID, "sysverb_insert").click()
        
        time.sleep(5) 
        return True 

    except Exception as e:
        print(f"❌ Error ServiceNow UAT: {e}")
        return False

# ==========================================
# Main: ระบบ Loop เฝ้าระบบ
# ==========================================
if __name__ == "__main__":
    db = initialize_firestore()
    if db:
        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222") # เชื่อมต่อ Chrome Debug
        try:
            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 15)
            print("🤖 บอทพร้อมทำงานบน UAT! (เฝ้าดูสถานะ on process)")

            while True:
                doc_id, doc_data = get_and_lock_ticket(db)
                if doc_id:
                    if fill_servicenow_ticket(driver, wait, doc_data):
                        mark_as_completed(db, doc_id)
                else:
                    print(".", end="", flush=True)
                time.sleep(15)
        except Exception as e:
            print(f"❌ ไม่พบ Chrome โหมด Debug: {e}")
            input("กด Enter เพื่อปิด...")
