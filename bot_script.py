import time
import os
import firebase_admin
from firebase_admin import credentials, firestore
from google.cloud.firestore_v1.base_query import FieldFilter
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
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
        if not firebase_admin._apps:
            firebase_admin.initialize_app(cred)
        return firestore.client()
    except Exception as e:
        print(f"❌ Error Firestore Init: {e}")
        return None

# --- สำหรับโหมดสร้าง Ticket ---
def get_and_lock_ticket(db):
    print("\n🔍 [Mode Create] ตรวจสอบงานใหม่ (on process)...")
    try:
        query = db.collection("incidents").where(filter=FieldFilter("status", "==", "on process")).where(filter=FieldFilter("is_pulled", "==", False)).limit(1)
        docs = query.get()
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

# --- สำหรับโหมด Sync Status (Monitor) ---
def get_sync_request_task(db):
    print("\n🔍 [Mode Sync] ตรวจสอบงานที่ขอ Sync (Request)...")
    try:
        query = db.collection("incidents").where(filter=FieldFilter("sync_status", "==", "Request")).limit(1)
        docs = query.get()
        for doc in docs:
            return doc.id, doc.to_dict()
        return None, None
    except Exception as e:
        print(f"❌ Error Firestore Sync: {e}")
        return None, None

def mark_as_completed(db, doc_id):
    try:
        db.collection("incidents").document(doc_id).update({"status": "Completed"})
        print(f"✅ [Firestore] ปรับสถานะเป็น Completed")
    except Exception as e:
        print(f"❌ Error Firestore Update: {e}")

# ==========================================
# ส่วนที่ 2: ควบคุม ServiceNow UAT
# ==========================================

# --- 2.1 ฟังก์ชันสร้าง Ticket (โค้ดเดิม 100%) ---
def fill_servicenow_ticket(driver, wait, data):
    try:
        print(f"🚀 เริ่มสร้าง Ticket: {data.get('description', 'No Subject')}")
        driver.switch_to.default_content()
        new_incident_url = "https://keristest.service-now.com/incident.do?sys_id=-1"
        driver.get(new_incident_url)
        print("⏳ รอหน้าฟอร์มกางตัว...")
        time.sleep(6) 

        try:
            if not driver.find_elements(By.ID, "incident.short_description"):
                wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "gsft_main")))
                print("📥 มุดเข้า iframe (gsft_main) สำเร็จ")
        except: pass

        print("- กรอก Short Description...")
        short_desc = wait.until(EC.presence_of_element_located((By.ID, "incident.short_description")))
        short_desc.clear()
        short_desc.send_keys(data.get("description", ""))

        print("- กรอก Caller...")
        caller_field = driver.find_element(By.ID, "sys_display.incident.caller_id")
        caller_field.clear()
        caller_field.send_keys(data.get("caller", ""), Keys.RETURN)
        time.sleep(2)

        print("- กรอก Assignment Group...")
        ag_field = driver.find_element(By.ID, "sys_display.incident.assignment_group")
        ag_field.clear()
        ag_field.send_keys("FTH Call Center", Keys.RETURN)
        time.sleep(2)

        print("- ตั้งค่า Category, Impact, Urgency...")
        try:
            cat_value = data.get("category", "Software") 
            Select(driver.find_element(By.ID, "incident.category")).select_by_visible_text(cat_value)
            Select(driver.find_element(By.ID, "incident.impact")).select_by_visible_text("5 - Minor")
            Select(driver.find_element(By.ID, "incident.urgency")).select_by_visible_text("5 - Minor")
        except Exception as e:
            print(f"  ⚠️ Dropdown Error: {e}")

        try:
            print("- สลับไป Tab External References...")
            tab_element = driver.find_element(By.XPATH, "//span[contains(text(), 'External References')]")
            driver.execute_script("arguments[0].click();", tab_element)
            time.sleep(1)
            ext_ref = driver.find_element(By.ID, "incident.u_extrefno4")
            ext_ref.clear()
            ext_ref.send_keys(data.get("ticket_id", ""))
        except Exception as e:
            print(f"  ⚠️ External Ref Error: {e}")

        print("💾 กำลังกดบันทึก...")
        driver.find_element(By.ID, "sysverb_insert").click()
        time.sleep(5) 
        return True 
    except Exception as e:
        print(f"❌ Error Create Flow: {e}")
        return False

# --- 2.2 ฟังก์ชัน Monitor (Sync Status) ---
def check_sync_status(driver, wait, data):
    try:
        ticket_id = data.get("ticket_id")
        print(f"🔍 [Monitor] ตรวจสอบสถานะ INC: {ticket_id}")
        
        # ไปหน้า List
        driver.get("https://keristest.service-now.com/now/nav/ui/classic/params/target/incident_list.do%3Fsysparm_query%3Dactive%3Dtrue")
        time.sleep(5)

        driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "gsft_main")))

        # 1. เลือก Dropdown เป็น "for text" (ใช้ Wildcard Selector เพื่อแก้ ID สุ่ม)
        print("- เลือกโหมด 'for text'...")
        search_select = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "select[id$='_select']")))
        Select(search_select).select_by_value("zztextsearchyy")
        time.sleep(1)

        # 2. กรอกเลข INC และกด Enter
        print(f"- ค้นหาเลข: {ticket_id}")
        search_input = driver.find_element(By.CSS_SELECTOR, "input[id$='_text']")
        search_input.clear()
        search_input.send_keys(ticket_id)
        search_input.send_keys(Keys.ENTER)
        
        time.sleep(5) # รอผลลัพธ์โหลด

        # 3. ตรวจสอบสถานะในคอลัมน์ State
        try:
            state_cell = driver.find_element(By.XPATH, "//tr[contains(@class, 'list_row')][1]/td[contains(@data-column, 'state') or contains(@aria-label, 'State')]")
            current_state = state_cell.text.strip()
            print(f"📍 สถานะใน ServiceNow: {current_state}")
            return current_state
        except:
            print("⚠️ ไม่พบ Record")
            return "Not Found"

    except Exception as e:
        print(f"❌ Error Sync Flow: {e}")
        return "Error"

# ==========================================
# Main Loop (สลับโหมดทำงาน)
# ==========================================
if __name__ == "__main__":
    db = initialize_firestore()
    if db:
        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        try:
            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 15)
            print("🤖 บอท UAT พร้อมทำงาน! (Create & Sync Mode)")

            while True:
                # ภารกิจที่ 1: สร้าง Ticket (on process)
                doc_id, doc_data = get_and_lock_ticket(db)
                if doc_id:
                    if fill_servicenow_ticket(driver, wait, doc_data):
                        mark_as_completed(db, doc_id)

                # ภารกิจที่ 2: ตรวจสอบสถานะ (sync_status == Request)
                s_id, s_data = get_sync_request_task(db)
                if s_id:
                    res = check_sync_status(driver, wait, s_data)
                    if res == "Resolved":
                        db.collection("incidents").document(s_id).update({
                            "status": "Completed",
                            "sync_status": "Done"
                        })
                        print(f"🎉 Sync สำเร็จ: {s_id} ปรับเป็น Done")
                    else:
                        print(f"⏳ {s_id} สถานะยังไม่ใช่ Resolved (คือ {res}) - ข้ามไป")

                print(".", end="", flush=True)
                time.sleep(15)
        except Exception as e:
            print(f"❌ ไม่พบ Chrome โหมด Debug: {e}")
            input("กด Enter เพื่อปิด...")
