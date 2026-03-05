import time
import os
import firebase_admin
from firebase_admin import credentials, firestore
from google.cloud.firestore_v1.base_query import FieldFilter 
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==========================================
# ส่วนที่ 1: จัดการ Firestore (แก้ Warning)
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
# ส่วนที่ 2: ควบคุม ServiceNow UAT
# ==========================================

# --- 2.1 โหมดตรวจสอบสถานะ (Sync Status) ---
def sync_status_mode(driver, wait, data):
    try:
        ticket_id = data.get("ticket_id")
        print(f"🔍 [Sync] ตรวจสอบ Ticket: {ticket_id}")
        
        driver.switch_to.window(driver.window_handles[-1])
        driver.get("https://keristest.service-now.com/now/nav/ui/classic/params/target/incident_list.do%3Fsysparm_query%3Dactive%3Dtrue")
        time.sleep(5)

        driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "gsft_main")))

        # 1. จัดการปุ่ม Filter (เช็คก่อนกด)
        filter_btn = wait.until(EC.presence_of_element_located((By.ID, "incident_filter_toggle_image")))
        is_expanded = filter_btn.get_attribute("aria-expanded")
        if is_expanded == "false":
            print("🖱️ กำลังคลิกเปิด Filter...")
            filter_btn.click()
            time.sleep(3)
        else:
            print("ℹ️ Filter เปิดอยู่แล้ว")

        # 2. จัดการ Select2 (ช่องแรก: Field Selection)
        print("- ตั้งค่าช่องแรก: External Ref No 4...")
        # คลิกที่กล่อง Select2 (มักจะเป็นตัวแรกในตาราง filter)
        s2_container = driver.find_element(By.CSS_SELECTOR, "div.select2-container")
        s2_container.click()
        time.sleep(1)
        
        # พิมพ์ค้นหาและกด Enter
        s2_input = wait.until(EC.presence_of_element_located((By.ID, "s2id_autogen2_search")))
        s2_input.send_keys("External Ref No 4")
        time.sleep(1)
        s2_input.send_keys(Keys.RETURN)
        time.sleep(2)

        # 3. เลือก contains (Operator)
        print("- ตั้งค่าช่องสอง: contains...")
        op_sel = Select(driver.find_element(By.CSS_SELECTOR, "select.condOperator"))
        op_sel.select_by_value("LIKE") # LIKE คือ contains

        # 4. กรอกเลข Ticket
        print(f"- กรอกเลขค้นหา: {ticket_id}")
        val_input = driver.find_element(By.CSS_SELECTOR, "input.filerTableInput")
        val_input.clear()
        val_input.send_keys(ticket_id)

        # 5. กด Run
        print("🖱️ กด Run...")
        run_btn = driver.find_element(By.ID, "test_filter_action_toolbar_run")
        driver.execute_script("arguments[0].click();", run_btn)
        time.sleep(4)

        # 6. เช็คสถานะ State
        try:
            state_cell = driver.find_element(By.XPATH, "//tr[contains(@class, 'list_row')][1]/td[contains(@data-column, 'state') or contains(@aria-label, 'State')]")
            current_state = state_cell.text.strip()
            print(f"📍 พบสถานะ: {current_state}")
            return current_state
        except:
            print("⚠️ ไม่พบข้อมูล")
            return "Not Found"

    except Exception as e:
        print(f"❌ Error Sync Mode: {e}")
        return "Error"

# --- 2.2 โหมดสร้าง Ticket ---
def create_ticket_mode(driver, wait, data):
    try:
        print(f"🚀 [สร้าง] เริ่มงาน: {data.get('description')}")
        driver.switch_to.default_content()
        driver.get("https://keristest.service-now.com/incident.do?sys_id=-1")
        time.sleep(5)
        
        try:
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "gsft_main")))
        except: pass

        wait.until(EC.presence_of_element_located((By.ID, "incident.short_description"))).send_keys(data.get("description", ""))
        driver.find_element(By.ID, "sys_display.incident.caller_id").send_keys(data.get("caller", ""), Keys.RETURN)
        time.sleep(1)
        driver.find_element(By.ID, "sys_display.incident.assignment_group").send_keys("FTH Call Center", Keys.RETURN)
        time.sleep(1)

        Select(driver.find_element(By.ID, "incident.category")).select_by_visible_text(data.get("category", "Software"))
        Select(driver.find_element(By.ID, "incident.impact")).select_by_visible_text("5 - Minor")
        Select(driver.find_element(By.ID, "incident.urgency")).select_by_visible_text("5 - Minor")

        tab = driver.find_element(By.XPATH, "//span[contains(text(), 'External References')]")
        driver.execute_script("arguments[0].click();", tab)
        time.sleep(1)
        driver.find_element(By.ID, "incident.u_extrefno4").send_keys(data.get("ticket_id", ""))

        driver.find_element(By.ID, "sysverb_insert").click()
        print("✅ สร้างสำเร็จ")
        return True
    except Exception as e:
        print(f"❌ Error Create Mode: {e}")
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
            print("🤖 บอท Master พร้อมรัน! (แก้ไข Iframe & Select2)")

            while True:
                # ภารกิจสร้าง
                c_id, c_data = get_create_task(db)
                if c_id:
                    if create_ticket_mode(driver, wait, c_data):
                        db.collection("incidents").document(c_id).update({"status": "Completed"})

                # ภารกิจ Sync
                s_id, s_data = get_sync_request_task(db)
                if s_id:
                    res = sync_status_mode(driver, wait, s_data)
                    if res == "Resolved":
                        db.collection("incidents").document(s_id).update({
                            "status": "Completed",
                            "sync_status": "Done"
                        })
                        print(f"🎉 Sync {s_id} สำเร็จ!")

                print(".", end="", flush=True)
                time.sleep(15)
        except Exception as e:
            print(f"❌ Error Main: {e}")
            input("กด Enter เพื่อปิด...")
