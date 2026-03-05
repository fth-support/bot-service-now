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

# ==========================================
# ส่วนที่ 1: จัดการ Firestore (ใช้ FieldFilter แก้ Warning)
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

# --- 2.1 โหมดสร้าง Ticket (ดึงกลับมาครบถ้วนทุกบรรทัด) ---
def create_ticket_mode(driver, wait, data):
    try:
        print(f"🚀 [โหมดสร้าง] เริ่มสร้าง Ticket: {data.get('ticket_id')}")
        driver.switch_to.default_content()
        driver.get("https://keristest.service-now.com/incident.do?sys_id=-1")
        time.sleep(5)
        
        # มุดเข้า Frame
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "gsft_main")))

        # กรอกข้อมูลพื้นฐาน
        wait.until(EC.presence_of_element_located((By.ID, "incident.short_description"))).send_keys(data.get("description", ""))
        driver.find_element(By.ID, "sys_display.incident.caller_id").send_keys(data.get("caller", ""), Keys.RETURN)
        time.sleep(1)
        driver.find_element(By.ID, "sys_display.incident.assignment_group").send_keys("FTH Call Center", Keys.RETURN)
        time.sleep(1)

        # เลือก Category, Impact, Urgency
        Select(driver.find_element(By.ID, "incident.category")).select_by_visible_text(data.get("category", "Software"))
        Select(driver.find_element(By.ID, "incident.impact")).select_by_visible_text("5 - Minor")
        Select(driver.find_element(By.ID, "incident.urgency")).select_by_visible_text("5 - Minor")

        # สลับไป Tab External References
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
        return False

# --- 2.2 โหมดตรวจสอบสถานะ (Sync Status) - ท่าใหม่ Search "for text" ---
def sync_status_mode(driver, wait, data):
    try:
        ticket_id = data.get("ticket_id")
        print(f"🔍 [โหมด Sync] ค้นหาด้วย Search for text: {ticket_id}")
        
        # เข้าหน้า List
        driver.get("https://keristest.service-now.com/now/nav/ui/classic/params/target/incident_list.do%3Fsysparm_query%3Dactive%3Dtrue")
        time.sleep(5)

        driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "gsft_main")))

        # 1. เลือก Dropdown เป็น "for text" (zztextsearchyy)
        print("- ตั้งค่าการค้นหาเป็น 'for text'...")
        # ใช้ CSS Selector ที่หา ID ลงท้ายด้วย _select เพื่อความยืดหยุ่น
        search_dropdown = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "select[id$='_select']")))
        Select(search_dropdown).select_by_value("zztextsearchyy")

        # 2. กรอก Ticket ID ในช่อง Search และกด Enter
        print(f"- กรอกเลขค้นหา: {ticket_id}")
        search_input = driver.find_element(By.CSS_SELECTOR, "input[id$='_text']")
        search_input.clear()
        search_input.send_keys(ticket_id)
        search_input.send_keys(Keys.RETURN)
        
        time.sleep(5) # รอผลลัพธ์แสดง

        # 3. เช็คสถานะ State ในแถวแรกของตาราง
        try:
            # ค้นหา Cell ในคอลัมน์ State
            state_cell = driver.find_element(By.XPATH, "//tr[contains(@class, 'list_row')][1]/td[contains(@data-column, 'state') or contains(@aria-label, 'State')]")
            current_state = state_cell.text.strip()
            print(f"📍 สถานะปัจจุบัน: {current_state}")
            return current_state
        except:
            print("⚠️ ไม่พบ Record จากการค้นหา")
            return "Not Found"

    except Exception as e:
        print(f"❌ Error Sync: {e}")
        return "Error"

# ==========================================
# Main Loop (สลับทำงาน 2 โหมด)
# ==========================================
if __name__ == "__main__":
    db = initialize_firestore()
    if db:
        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        try:
            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 15)
            print("🤖 บอท Master พร้อมรัน! (โหมด Search for text ลื่นไหลกว่าเดิม)")

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
                        print(f"🎉 Sync สำเร็จ: {s_id}")

                print(".", end="", flush=True)
                time.sleep(15)
        except Exception as e:
            print(f"❌ Error Main: {e}")
            input("กด Enter เพื่อปิด...")
