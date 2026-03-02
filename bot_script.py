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
# ส่วนที่ 1: จัดการ Firestore
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
    print("\n🔍 ตรวจสอบงานใหม่ (on process)...")
    try:
        docs = db.collection("incidents") \
                 .where("status", "==", "on process") \
                 .where("is_pulled", "==", False) \
                 .limit(1).get()
        
        for doc in docs:
            doc_id = doc.id
            doc_data = doc.to_dict()
            db.collection("incidents").document(doc_id).update({"is_pulled": True})
            print(f"🔒 ล็อกงาน {doc_id} เรียบร้อย")
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
# ส่วนที่ 2: ควบคุม ServiceNow UAT
# ==========================================
def fill_servicenow_ticket(driver, wait, data):
    try:
        print(f"🚀 เริ่มสร้าง Ticket: {data.get('description', 'No Subject')}")
        
        # --- ⚡ ท่าไม้ตาย: พุ่งตรงไปหน้า New Form (ตัดปัญหาปุ่ม New หาย) ⚡ ---
        # ใช้ URL ที่เข้าหน้าสร้าง Incident ใหม่โดยตรง (Direct URL)
        new_incident_url = "https://keristest.service-now.com/incident.do?sys_id=-1"
        driver.get(new_incident_url)
        
        print("⏳ กำลังโหลดหน้าฟอร์มใหม่...")
        time.sleep(5) 

        # ตรวจสอบว่าต้องมุด iframe ไหม (กรณี ServiceNow ครอบ UI ไว้)
        driver.switch_to.default_content()
        try:
            # ลองหาฟิลด์ Short Description ถ้าไม่เจอแสดงว่าอาจติด iframe
            if not driver.find_elements(By.ID, "incident.short_description"):
                wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "gsft_main")))
                print("📥 มุดเข้า iframe (gsft_main) สำเร็จ")
        except:
            pass

        # 1. กรอก Short Description (จากฟิลด์ description)
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

        # 3. กรอก Assignment Group (FTH Call Center)
        print("- กรอก Assignment Group (FTH Call Center)...")
        ag_field = driver.find_element(By.ID, "sys_display.incident.assignment_group")
        ag_field.clear()
        ag_field.send_keys("FTH Call Center", Keys.RETURN)
        time.sleep(2)

        # 4. กรอก External Ref No 4 (ใช้ ID จริง: incident.u_extrefno4)
        print("- สลับไป Tab External References และกรอกข้อมูล...")
        try:
            # คลิก Tab ก่อน (หาจากชื่อ Tab)
            tab_element = driver.find_element(By.XPATH, "//span[contains(text(), 'External References')]")
            driver.execute_script("arguments[0].click();", tab_element)
            time.sleep(1)

            # กรอกข้อมูลลงช่อง ID: incident.u_extrefno4
            ext_ref = wait.until(EC.visibility_of_element_located((By.ID, "incident.u_extrefno4")))
            ext_ref.clear()
            ext_ref.send_keys(data.get("ticket_id", ""))
            print("  ✅ กรอก External Ref สำเร็จ")
        except Exception as ex:
            print(f"  ⚠️ หา Tab หรือช่อง External Ref ไม่เจอ: {ex}")

        # 5. กด Save (Submit)
        print("💾 กำลังกดบันทึก...")
        driver.find_element(By.ID, "sysverb_insert").click()
        
        time.sleep(5) 
        return True 

    except Exception as e:
        print(f"❌ Error ServiceNow Flow: {e}")
        return False

# ==========================================
# Main: ระบบ Loop
# ==========================================
if __name__ == "__main__":
    db = initialize_firestore()
    if db:
        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        try:
            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 15)
            print("🤖 บอทพร้อมรันระบบ UAT! (โหมด Direct Link)")

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
