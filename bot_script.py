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

def get_and_lock_ticket(db):
    try:
        docs = db.collection("incidents").where("status", "==", "on process").where("is_pulled", "==", False).limit(1).get()
        for doc in docs:
            doc_id = doc.id
            doc_data = doc.to_dict()
            db.collection("incidents").document(doc_id).update({"is_pulled": True})
            return doc_id, doc_data
        return None, None
    except Exception as e:
        print(f"❌ Error Firestore Create: {e}")
        return None, None

def get_sync_request_task(db):
    try:
        docs = db.collection("incidents").where("sync_status", "==", "Request").limit(1).get()
        for doc in docs:
            return doc.id, doc.to_dict()
        return None, None
    except Exception as e:
        print(f"❌ Error Firestore Sync: {e}")
        return None, None

def mark_as_completed(db, doc_id):
    try:
        db.collection("incidents").document(doc_id).update({"status": "Completed"})
    except: pass

# ==========================================
# ส่วนที่ 2: ควบคุม ServiceNow UAT
# ==========================================

# --- ท่าบังคับ: ดึงบอทกลับมา Tab ปัจจุบัน ---
def force_active_tab(driver):
    try:
        # บังคับให้บอททำงานที่หน้าต่างล่าสุดเสมอ เพื่อไม่ให้หน้าจอนิ่ง
        driver.switch_to.window(driver.window_handles[-1])
    except: pass

# --- 2.1 โหมดสร้าง Ticket ---
def fill_servicenow_ticket(driver, wait, data):
    try:
        print(f"🚀 [Create] เริ่มสร้าง Ticket: {data.get('ticket_id', 'Unknown')}")
        force_active_tab(driver)
        driver.get("https://keristest.service-now.com/incident.do?sys_id=-1")
        print("⏳ รอหน้าเว็บโหลด...")
        time.sleep(5) 

        # ⚡ ท่าแก้หน้าจอนิ่ง: เช็ค Iframe แบบไม่ต้องรอ (ถ้ามีก็มุด ไม่มีก็ข้าม)
        driver.switch_to.default_content()
        if driver.find_elements(By.ID, "gsft_main"):
            driver.switch_to.frame("gsft_main")
            print("📥 มุดเข้า iframe สำเร็จ")

        print("- กรอก Short Description และ Caller...")
        wait.until(EC.presence_of_element_located((By.ID, "incident.short_description"))).send_keys(data.get("description", ""))
        
        caller = driver.find_element(By.ID, "sys_display.incident.caller_id")
        caller.clear()
        caller.send_keys(data.get("caller", ""), Keys.RETURN)
        time.sleep(1)

        print("- กรอก Assignment Group...")
        ag = driver.find_element(By.ID, "sys_display.incident.assignment_group")
        ag.clear()
        ag.send_keys("FTH Call Center", Keys.RETURN)
        time.sleep(1)

        print("- ตั้งค่า Dropdowns...")
        Select(driver.find_element(By.ID, "incident.category")).select_by_visible_text(data.get("category", "Software"))
        Select(driver.find_element(By.ID, "incident.impact")).select_by_visible_text("5 - Minor")
        Select(driver.find_element(By.ID, "incident.urgency")).select_by_visible_text("5 - Minor")

        print("- สลับ Tab และกรอก External Ref...")
        tab = driver.find_element(By.XPATH, "//span[contains(text(), 'External References')]")
        driver.execute_script("arguments[0].click();", tab)
        time.sleep(1)
        driver.find_element(By.ID, "incident.u_extrefno4").clear()
        driver.find_element(By.ID, "incident.u_extrefno4").send_keys(data.get("ticket_id", ""))

        print("💾 กำลังกด Save...")
        driver.find_element(By.ID, "sysverb_insert").click()
        time.sleep(5) 
        print("✅ สร้าง Ticket สำเร็จ")
        return True 
    except Exception as e:
        print(f"❌ Error Create Flow: {e}")
        return False

# --- 2.2 โหมด Track/Monitor ---
def check_sync_status(driver, wait, data):
    try:
        ticket_id = data.get("ticket_id")
        print(f"🔍 [Track] เริ่มค้นหาสถานะของ: {ticket_id}")
        force_active_tab(driver)
        driver.get("https://keristest.service-now.com/incident_list.do")
        print("⏳ รอหน้าเว็บโหลด...")
        time.sleep(4)

        # ⚡ เช็ค Iframe แบบรวดเร็ว
        driver.switch_to.default_content()
        if driver.find_elements(By.ID, "gsft_main"):
            driver.switch_to.frame("gsft_main")

        # 1. เลือก Dropdown เป็น "for text" (ใช้ XPATH เล็ง ID ที่ลงท้ายด้วย _select)
        print("- เลือกโหมด 'for text'...")
        search_dropdown = wait.until(EC.presence_of_element_located((By.XPATH, "//select[contains(@id, '_select') and contains(@class, 'form-control')]")))
        Select(search_dropdown).select_by_value("zztextsearchyy")
        time.sleep(1)

        # 2. กรอก INC และกด Enter (ใช้ XPATH เล็ง ID ที่ลงท้ายด้วย _text)
        print(f"- พิมพ์ค้นหา: {ticket_id}")
        search_input = driver.find_element(By.XPATH, "//input[@type='search' and contains(@id, '_text')]")
        search_input.clear()
        search_input.send_keys(ticket_id)
        time.sleep(1)
        search_input.send_keys(Keys.ENTER)
        
        print("⏳ รอระบบโหลดผลการค้นหา...")
        time.sleep(4) 

        # 3. เช็คสถานะ Resolved จากข้อมูลในแถว
        try:
            # ดึง Text จากทั้งแถว (Row) มาเช็คเลยว่ามีคำว่า Resolved ไหม (ชัวร์และง่ายที่สุด)
            row = wait.until(EC.presence_of_element_located((By.XPATH, "//tr[contains(@class, 'list_row')][1]")))
            row_text = row.text
            print(f"📍 ข้อมูลที่พบในระบบ: {row_text}")
            
            if "Resolved" in row_text:
                return "Resolved"
            else:
                return "Open"
        except:
            print("⚠️ ไม่พบ Record (อาจยังไม่เข้าระบบ)")
            return "Not Found"

    except Exception as e:
        print(f"❌ Error Sync Flow: {e}")
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
            print("🤖 บอท UAT พร้อมทำงาน! (แก้ปัญหาหน้าจอนิ่ง 100%)")

            while True:
                # ภารกิจ 1: Create
                doc_id, doc_data = get_and_lock_ticket(db)
                if doc_id:
                    if fill_servicenow_ticket(driver, wait, doc_data):
                        mark_as_completed(db, doc_id)

                # ภารกิจ 2: Monitor
                s_id, s_data = get_sync_request_task(db)
                if s_id:
                    res = check_sync_status(driver, wait, s_data)
                    if res == "Resolved":
                        db.collection("incidents").document(s_id).update({
                            "status": "Completed",
                            "sync_status": "Done"
                        })
                        print(f"🎉 ตรวจพบ Resolved -> อัปเดต {s_id} สำเร็จ!")
                    else:
                        print(f"⏳ {s_id} สถานะยังไม่ใช่ Resolved (ข้ามไปก่อน)")

                print(".", end="", flush=True)
                time.sleep(15)
                
        except Exception as e:
            print(f"❌ ไม่พบ Chrome โหมด Debug: {e}")
            input("กด Enter เพื่อปิด...")
