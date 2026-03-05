import time
import os
import firebase_admin
from firebase_admin import credentials, firestore
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

# --- Logic สำหรับการสร้างงานใหม่ ---
def get_create_task(db):
    docs = db.collection("incidents").where("status", "==", "on process").where("is_pulled", "==", False).limit(1).get()
    for doc in docs:
        db.collection("incidents").document(doc.id).update({"is_pulled": True})
        return doc.id, doc.to_dict()
    return None, None

# --- Logic สำหรับการ Sync Status (ภารกิจใหม่) ---
def get_sync_request_task(db):
    docs = db.collection("incidents").where("sync_status", "==", "Request").limit(1).get()
    for doc in docs:
        return doc.id, doc.to_dict()
    return None, None

# ==========================================
# ส่วนที่ 2: ควบคุม ServiceNow UAT
# ==========================================

# 2.1 โหมดสร้าง Ticket (Logic เดิม)
def create_new_ticket(driver, wait, data):
    try:
        driver.switch_to.default_content()
        driver.get("https://keristest.service-now.com/incident.do?sys_id=-1")
        time.sleep(5)
        # (ส่วนการกรอกข้อมูลเหมือนเดิมที่คุณ POC ผ่านแล้ว)
        wait.until(EC.presence_of_element_located((By.ID, "incident.short_description"))).send_keys(data.get("description", ""))
        driver.find_element(By.ID, "sys_display.incident.caller_id").send_keys(data.get("caller", ""), Keys.RETURN)
        time.sleep(2)
        driver.find_element(By.ID, "sys_display.incident.assignment_group").send_keys("FTH Call Center", Keys.RETURN)
        
        # เลือก Category, Impact, Urgency
        Select(driver.find_element(By.ID, "incident.category")).select_by_visible_text(data.get("category", "Software"))
        Select(driver.find_element(By.ID, "incident.impact")).select_by_visible_text("5 - Minor")
        Select(driver.find_element(By.ID, "incident.urgency")).select_by_visible_text("5 - Minor")
        
        # External Ref
        tab_element = driver.find_element(By.XPATH, "//span[contains(text(), 'External References')]")
        driver.execute_script("arguments[0].click();", tab_element)
        time.sleep(1)
        driver.find_element(By.ID, "incident.u_extrefno4").send_keys(data.get("ticket_id", ""))
        
        driver.find_element(By.ID, "sysverb_insert").click()
        time.sleep(5)
        return True
    except Exception as e:
        print(f"❌ Error Create: {e}")
        return False

# 2.2 โหมด Sync Check (ภารกิจใหม่)
def check_ticket_sync_status(driver, wait, data):
    try:
        ticket_id = data.get("ticket_id")
        print(f"🔍 เริ่มตรวจสอบ Sync สำหรับ: {ticket_id}")
        
        # ไปหน้า List
        list_url = "https://keristest.service-now.com/now/nav/ui/classic/params/target/incident_list.do%3Fsysparm_query%3Dactive%3Dtrue"
        driver.get(list_url)
        time.sleep(5)
        
        driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "gsft_main")))

        # 1. กดปุ่ม Filter icon
        filter_btn = wait.until(EC.element_to_be_clickable((By.ID, "incident_filter_toggle_image")))
        filter_btn.click()
        time.sleep(1)

        # 2. เลือกช่องแรกเป็น "External Ref No 4" (ใช้ท่า Select2 ทะลวง)
        print("- ตั้งค่า Filter: External Ref No 4...")
        # คลิกที่ตัวเลือกแรก (ปกติคือ Number หรือช่องแรก)
        first_select = driver.find_element(By.CSS_SELECTOR, "select.filerTableSelect")
        Select(first_select).select_by_visible_text("External Ref No 4")
        
        # 3. เลือกช่องสองเป็น "contains"
        operator_select = driver.find_element(By.CSS_SELECTOR, "select.condOperator")
        Select(operator_select).select_by_visible_text("contains")

        # 4. กรอก ticket_id ในช่องสุดท้าย
        val_input = driver.find_element(By.CSS_SELECTOR, "input.filerTableInput")
        val_input.clear()
        val_input.send_keys(ticket_id)

        # 5. กด Run
        print("- กดปุ่ม Run Filter...")
        driver.find_element(By.ID, "test_filter_action_toolbar_run").click()
        time.sleep(4)

        # 6. ตรวจสอบสถานะ State ในตาราง
        try:
            # ค้นหา Cell ในคอลัมน์ State (ปกติ ServiceNow จะมี aria-label หรือ class)
            # เราจะหาแถวแรกที่โผล่มา
            state_cell = driver.find_element(By.XPATH, "//tr[@class='list_row']/td[contains(@data-column, 'state') or contains(@aria-label, 'State')]")
            current_state = state_cell.text.strip()
            print(f"📍 สถานะปัจจุบันใน ServiceNow: {current_state}")

            if current_state == "Resolved":
                return "Resolved"
            else:
                return "Open"
        except:
            print("⚠️ ไม่พบ Record จากการ Filter")
            return "Not Found"

    except Exception as e:
        print(f"❌ Error Sync Check: {e}")
        return "Error"

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
            print("🤖 บอท UAT พร้อมทำงาน (โหมด Dual: Create & Sync Status)")

            while True:
                # --- งานที่ 1: สร้าง Ticket ใหม่ ---
                c_id, c_data = get_create_task(db)
                if c_id:
                    if create_new_ticket(driver, wait, c_data):
                        db.collection("incidents").document(c_id).update({"status": "Completed"})
                        print(f"✅ สร้าง Ticket {c_id} สำเร็จ")
                
                # --- งานที่ 2: Sync Check (Request) ---
                s_id, s_data = get_sync_request_task(db)
                if s_id:
                    result = check_ticket_sync_status(driver, wait, s_data)
                    if result == "Resolved":
                        db.collection("incidents").document(s_id).update({
                            "status": "Completed",
                            "sync_status": "Done"
                        })
                        print(f"🎉 Sync สำเร็จ: {s_id} ถูกปรับเป็น Completed/Done")
                    else:
                        print(f"ℹ️ {s_id} ยังไม่ Resolved (สถานะ: {result}) ไม่มีการอัปเดต Firebase")

                print(".", end="", flush=True)
                time.sleep(15)
        except Exception as e:
            print(f"❌ Error Main: {e}")
            input("Press Enter to exit...")
