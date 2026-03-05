import time
import os
import firebase_admin
from firebase_admin import credentials, firestore
from google.cloud.firestore_v1.base_query import FieldFilter # แก้ Warning
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select, WebDriverWait
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

def get_create_task(db):
    # ค้นหางานสร้างใหม่โดยใช้ FieldFilter เพื่อความเสถียร
    query = db.collection("incidents").where(filter=FieldFilter("status", "==", "on process")).where(filter=FieldFilter("is_pulled", "==", False)).limit(1)
    docs = query.get()
    for doc in docs:
        db.collection("incidents").document(doc.id).update({"is_pulled": True})
        return doc.id, doc.to_dict()
    return None, None

def get_sync_request_task(db):
    # ค้นหางานตรวจสอบสถานะ
    query = db.collection("incidents").where(filter=FieldFilter("sync_status", "==", "Request")).limit(1)
    docs = query.get()
    for doc in docs:
        return doc.id, doc.to_dict()
    return None, None

# ==========================================
# ส่วนที่ 2: ฟังก์ชันควบคุม ServiceNow (Manual UI)
# ==========================================

# --- 2.1 โหมดสร้าง Ticket (Full UI) ---
def create_ticket_mode(driver, wait, data):
    try:
        print(f"🚀 [สร้าง] เริ่มงาน: {data.get('ticket_id')}")
        driver.switch_to.default_content()
        driver.get("https://keristest.service-now.com/incident.do?sys_id=-1")
        time.sleep(5)
        
        try:
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "gsft_main")))
        except: pass

        # กรอกข้อมูล Short Description
        wait.until(EC.presence_of_element_located((By.ID, "incident.short_description"))).send_keys(data.get("description", ""))
        
        # กรอก Caller และ Assignment Group
        driver.find_element(By.ID, "sys_display.incident.caller_id").send_keys(data.get("caller", ""), Keys.RETURN)
        time.sleep(1)
        driver.find_element(By.ID, "sys_display.incident.assignment_group").send_keys("FTH Call Center", Keys.RETURN)
        time.sleep(1)

        # เลือก Dropdowns ตาม HTML จริง
        Select(driver.find_element(By.ID, "incident.category")).select_by_visible_text(data.get("category", "Software"))
        Select(driver.find_element(By.ID, "incident.impact")).select_by_visible_text("5 - Minor")
        Select(driver.find_element(By.ID, "incident.urgency")).select_by_visible_text("5 - Minor")

        # สลับ Tab External References
        tab = driver.find_element(By.XPATH, "//span[contains(text(), 'External References')]")
        driver.execute_script("arguments[0].click();", tab)
        time.sleep(1)
        driver.find_element(By.ID, "incident.u_extrefno4").send_keys(data.get("ticket_id", ""))

        # กดปุ่มบันทึก
        driver.find_element(By.ID, "sysverb_insert").click()
        print("✅ สร้าง Ticket สำเร็จ")
        return True
    except Exception as e:
        print(f"❌ Error สร้าง: {e}")
        return False

# --- 2.2 โหมด Sync Status (Full UI - เห็นการกด Filter) ---
def sync_status_mode(driver, wait, data):
    try:
        ticket_id = data.get("ticket_id")
        print(f"🔍 [Sync] ตรวจสอบ Ticket: {ticket_id}")
        
        # บังคับ Focus หน้าต่างและเข้าหน้า List
        driver.switch_to.window(driver.window_handles[-1])
        driver.get("https://keristest.service-now.com/now/nav/ui/classic/params/target/incident_list.do%3Fsysparm_query%3Dactive%3Dtrue")
        time.sleep(5)

        driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "gsft_main")))

        # 1. คลิกปุ่ม Filter (นี่คือจุดที่โค้ดเพิ่มขึ้นมา เพื่อให้เห็นหน้าจอขยับ)
        print("🖱️ กำลังคลิกปุ่ม Filter...")
        wait.until(EC.element_to_be_clickable((By.ID, "incident_filter_toggle_image"))).click()
        time.sleep(2)

        # 2. ตั้งค่าเงื่อนไข Filter ทีละช่อง
        print("- เลือก External Ref No 4...")
        Select(wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "select.filerTableSelect")))).select_by_visible_text("External Ref No 4")
        
        print("- เลือก contains...")
        Select(driver.find_element(By.CSS_SELECTOR, "select.condOperator")).select_by_visible_text("contains")
        
        print(f"- พิมพ์เลข: {ticket_id}")
        val_input = driver.find_element(By.CSS_SELECTOR, "input.filerTableInput")
        val_input.clear()
        val_input.send_keys(ticket_id)

        # 3. กดปุ่ม Run
        print("🖱️ กำลังกดปุ่ม Run Filter...")
        driver.find_element(By.ID, "test_filter_action_toolbar_run").click()
        time.sleep(4)

        # 4. ตรวจสอบสถานะจากตารางผลลัพธ์
        try:
            state_text = driver.find_element(By.XPATH, "//tr[contains(@class, 'list_row')][1]/td[contains(@data-column, 'state') or contains(@aria-label, 'State')]").text
            print(f"📍 สถานะปัจจุบัน: {state_text}")
            return state_text.strip()
        except:
            print("⚠️ ไม่พบแถวข้อมูล")
            return "Not Found"

    except Exception as e:
        print(f"❌ Error Sync: {e}")
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
            print("🤖 บอท Master พร้อมรัน! (โหมด Full UI ขยับทุกสเต็ป)")

            while True:
                # ภารกิจสร้าง Ticket
                c_id, c_data = get_create_task(db)
                if c_id:
                    if create_ticket_mode(driver, wait, c_data):
                        db.collection("incidents").document(c_id).update({"status": "Completed"})

                # ภารกิจ Sync Status
                s_id, s_data = get_sync_request_task(db)
                if s_id:
                    res = sync_status_mode(driver, wait, s_data)
                    if res == "Resolved":
                        db.collection("incidents").document(s_id).update({
                            "status": "Completed",
                            "sync_status": "Done"
                        })
                        print(f"🎉 Sync สำเร็จ: {s_id} เป็น Done")

                print(".", end="", flush=True)
                time.sleep(15)
        except Exception as e:
            print(f"❌ Error Main: {e}")
            input("กด Enter เพื่อปิด...")
