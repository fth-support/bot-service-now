import time
import os
from datetime import datetime
import firebase_admin
from firebase_admin import credentials, firestore
from google.cloud.firestore_v1.base_query import FieldFilter 
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- 🛠️ ระบบ Debug Tool (ถ่ายรูปเมื่อพลาด) ---
def save_debug_info(driver, task_id, phase):
    if not os.path.exists("debug_logs"):
        os.makedirs("debug_logs")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"debug_logs/{task_id}_{phase}_{timestamp}.png"
    driver.save_screenshot(filename)
    print(f"📸 บันทึกภาพหน้าจอไว้ที่: {filename}")

# ==========================================
# ส่วนที่ 1: จัดการ Firestore (แก้ Warning เรื่อง filter=)
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

def get_create_task(db):
    query = db.collection("incidents").where(filter=FieldFilter("status", "==", "on process")).where(filter=FieldFilter("is_pulled", "==", False)).limit(1)
    docs = query.get()
    for doc in docs:
        db.collection("incidents").document(doc.id).update({"is_pulled": True})
        return doc.id, doc.to_dict()
    return None, None

def get_sync_request_task(db):
    query = db.collection("incidents").where(filter=FieldFilter("sync_status", "==", "Request")).limit(1)
    docs = query.get()
    for doc in docs:
        return doc.id, doc.to_dict()
    return None, None

# ==========================================
# ส่วนที่ 2: ฟังก์ชันควบคุม ServiceNow
# ==========================================

# --- 2.1 โหมดสร้าง Ticket (ดึงโค้ดที่เคยหายไปกลับมาครบถ้วน) ---
def create_ticket_mode(driver, wait, data, task_id):
    try:
        print(f"🚀 [โหมดสร้าง] เริ่มสร้าง Ticket: {data.get('ticket_id')}")
        driver.switch_to.default_content()
        driver.get("https://keristest.service-now.com/incident.do?sys_id=-1")
        time.sleep(5)
        
        # มุดเข้า Frame gsft_main
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "gsft_main")))

        # กรอกข้อมูล Short Description และ Caller
        wait.until(EC.presence_of_element_located((By.ID, "incident.short_description"))).send_keys(data.get("description", ""))
        driver.find_element(By.ID, "sys_display.incident.caller_id").send_keys(data.get("caller", ""), Keys.RETURN)
        time.sleep(1)

        # กรอก Assignment Group
        ag_field = driver.find_element(By.ID, "sys_display.incident.assignment_group")
        ag_field.clear()
        ag_field.send_keys("FTH Call Center", Keys.RETURN)
        time.sleep(1)

        # เลือก Category, Impact, Urgency เป็น 5 - Minor
        Select(driver.find_element(By.ID, "incident.category")).select_by_visible_text(data.get("category", "Software"))
        Select(driver.find_element(By.ID, "incident.impact")).select_by_visible_text("5 - Minor")
        Select(driver.find_element(By.ID, "incident.urgency")).select_by_visible_text("5 - Minor")

        # สลับไป Tab External References และกรอกข้อมูล
        tab = driver.find_element(By.XPATH, "//span[contains(text(), 'External References')]")
        driver.execute_script("arguments[0].click();", tab)
        time.sleep(1)
        driver.find_element(By.ID, "incident.u_extrefno4").send_keys(data.get("ticket_id", ""))

        # กด Submit
        driver.find_element(By.ID, "sysverb_insert").click()
        print("✅ สร้าง Ticket สำเร็จ")
        return True
    except Exception as e:
        print(f"❌ Error สร้าง: {e}")
        save_debug_info(driver, task_id, "create_error")
        return False

# --- 2.2 โหมดตรวจสอบสถานะ (Sync Status) ---
def sync_status_mode(driver, wait, data, task_id):
    try:
        ticket_id = data.get("ticket_id")
        print(f"🔍 [โหมด Sync] ตรวจสอบ Ticket: {ticket_id}")
        
        driver.switch_to.window(driver.window_handles[-1])
        driver.get("https://keristest.service-now.com/now/nav/ui/classic/params/target/incident_list.do%3Fsysparm_query%3Dactive%3Dtrue")
        time.sleep(5)

        driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "gsft_main")))

        # 1. คลิกปุ่ม Filter (เช็คก่อนกด)
        filter_btn = wait.until(EC.presence_of_element_located((By.ID, "incident_filter_toggle_image")))
        if filter_btn.get_attribute("aria-expanded") == "false":
            print("🖱️ คลิกเปิด Filter...")
            filter_btn.click()
            time.sleep(3) # จังหวะรอโหลดฟอร์ม 2.8 วินาทีตาม HAR

        # 2. จัดการ Select2 สำหรับ External Ref No 4
        print("- เลือกฟิลด์: External Ref No 4...")
        s2_trigger = driver.find_element(By.CSS_SELECTOR, "a.select2-choice")
        driver.execute_script("arguments[0].click();", s2_trigger)
        time.sleep(1)
        
        # ค้นหาใน Select2
        s2_input = wait.until(EC.visibility_of_element_located((By.ID, "s2id_autogen2_search")))
        s2_input.send_keys("External Ref No 4")
        time.sleep(1)
        s2_input.send_keys(Keys.RETURN)
        time.sleep(1)

        # 3. เลือก contains (Operator)
        Select(driver.find_element(By.CSS_SELECTOR, "select.condOperator")).select_by_value("LIKE")

        # 4. กรอกเลข INC และกด Run
        val_input = driver.find_element(By.CSS_SELECTOR, "input.filerTableInput")
        val_input.clear()
        val_input.send_keys(ticket_id)
        
        print("🖱️ กด Run Filter...")
        driver.find_element(By.ID, "test_filter_action_toolbar_run").click()
        time.sleep(4)

        # 5. เช็คสถานะ State ในแถวแรก
        try:
            state_cell = driver.find_element(By.XPATH, "//tr[contains(@class, 'list_row')][1]/td[contains(@data-column, 'state') or contains(@aria-label, 'State')]")
            status_text = state_cell.text.strip()
            print(f"📍 สถานะปัจจุบัน: {status_text}")
            return status_text
        except:
            print("⚠️ ไม่พบแถวข้อมูล")
            return "Not Found"

    except Exception as e:
        print(f"❌ Error Sync: {e}")
        save_debug_info(driver, task_id, "sync_error")
        return "Error"

# ==========================================
# Main Loop (Dual Mode)
# ==========================================
if __name__ == "__main__":
    db = initialize_firestore()
    if db:
        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        try:
            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 15)
            print("🤖 บอท Master พร้อมรัน! (Full Code - Create & Sync)")

            while True:
                # ภารกิจสร้าง Ticket
                c_id, c_data = get_create_task(db)
                if c_id:
                    if create_ticket_mode(driver, wait, c_data, c_id):
                        db.collection("incidents").document(c_id).update({"status": "Completed"})

                # ภารกิจ Sync Status (Request)
                s_id, s_data = get_sync_request_task(db)
                if s_id:
                    res = sync_status_mode(driver, wait, s_data, s_id)
                    if res == "Resolved":
                        db.collection("incidents").document(s_id).update({
                            "status": "Completed",
                            "sync_status": "Done"
                        })
                        print(f"🎉 Sync สำเร็จ: {s_id}")

                print(".", end="", flush=True)
                time.sleep(15)
        except Exception as e:
            print(f"❌ Error Main: {e}")
            input("กด Enter เพื่อปิด...")
