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
        docs = db.collection("incidents") \
                 .where("status", "==", "on process") \
                 .where("is_pulled", "==", False) \
                 .limit(1).get()
        
        for doc in docs:
            doc_id = doc.id
            doc_data = doc.to_dict()
            db.collection("incidents").document(doc_id).update({"is_pulled": True})
            print(f"🔒 ล็อกงาน Create: {doc_id} เรียบร้อย")
            return doc_id, doc_data
        return None, None
    except Exception as e:
        print(f"❌ Error Firestore Create: {e}")
        return None, None

def get_sync_request_task(db):
    try:
        # บอทจะดึงมาทำเฉพาะงานที่เป็น "Request" เท่านั้น
        docs = db.collection("incidents") \
                 .where("sync_status", "==", "Request") \
                 .limit(1).get()
                 
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

def force_active_tab(driver):
    try:
        driver.switch_to.window(driver.window_handles[-1])
    except: pass

# --- งานที่ 2.1: สร้าง Ticket ---
def fill_servicenow_ticket(driver, wait, data):
    try:
        print(f"🚀 เริ่มสร้าง Ticket: {data.get('description', 'No Subject')}")
        
        driver.switch_to.default_content()
        new_incident_url = "https://keristest.service-now.com/incident.do?sys_id=-1"
        driver.get(new_incident_url)
        
        print("⏳ รอหน้าฟอร์มกางตัว...")
        time.sleep(6) 

        driver.switch_to.default_content()
        if driver.find_elements(By.ID, "gsft_main"):
            driver.switch_to.frame("gsft_main")
            print("📥 มุดเข้า iframe (gsft_main) สำเร็จ")

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

        print("- กำลังตั้งค่า Category, Impact, Urgency...")
        try:
            cat_value = data.get("category", "Software") 
            Select(driver.find_element(By.ID, "incident.category")).select_by_visible_text(cat_value)
            Select(driver.find_element(By.ID, "incident.impact")).select_by_visible_text("5 - Minor")
            Select(driver.find_element(By.ID, "incident.urgency")).select_by_visible_text("5 - Minor")
            print("  ✅ ตั้งค่า Dropdowns สำเร็จ")
        except Exception as e:
            print(f"  ⚠️ ไม่สามารถเลือก Dropdown ได้: {e}")

        try:
            print("- สลับไป Tab External References...")
            tab_element = driver.find_element(By.XPATH, "//span[contains(text(), 'External References')]")
            driver.execute_script("arguments[0].click();", tab_element)
            time.sleep(1)

            print("- กรอก External Ref No 4...")
            ext_ref = driver.find_element(By.ID, "incident.u_extrefno4")
            ext_ref.clear()
            ext_ref.send_keys(data.get("ticket_id", ""))
            print("  ✅ กรอก External Ref สำเร็จ")
        except Exception as e:
            print(f"  ⚠️ หา Tab หรือช่อง External Ref ไม่เจอ: {e}")

        print("💾 กำลังกดบันทึก...")
        driver.find_element(By.ID, "sysverb_insert").click()
        
        time.sleep(5) 
        return True 

    except Exception as e:
        print(f"❌ Error ServiceNow Flow: {e}")
        return False

# --- งานที่ 2.2: Track / Monitor Status ---
def check_sync_status(driver, wait, data):
    try:
        ticket_id = data.get("ticket_id")
        print(f"🔍 [Monitor] ค้นหาสถานะของ: {ticket_id}")
        
        force_active_tab(driver)
        driver.switch_to.default_content()
        driver.get("https://keristest.service-now.com/incident_list.do")
        time.sleep(5)

        try:
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "gsft_main")))
        except:
            pass

        print("- เลือกโหมด 'for text'...")
        search_dropdown_element = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(@class, 'input-group-select')]//select")))
        Select(search_dropdown_element).select_by_value("zztextsearchyy")
        time.sleep(1)

        print(f"- กรอกเลขค้นหา: {ticket_id}")
        search_input = driver.find_element(By.XPATH, "//div[contains(@class, 'input-group')]//input[@type='search']")
        search_input.clear()
        search_input.send_keys(ticket_id)
        time.sleep(1)
        search_input.send_keys(Keys.ENTER)
        
        print("⏳ รอระบบโหลดผลลัพธ์...")
        time.sleep(5) 

        try:
            state_cell = driver.find_element(By.XPATH, "//tr[contains(@class, 'list_row')][1]/td[contains(@data-column, 'state') or contains(@aria-label, 'State')]")
            current_state = state_cell.text.strip()
            print(f"📍 สถานะปัจจุบัน: {current_state}")
            return current_state
        except:
            print("⚠️ ไม่พบ Record (อาจค้นหาไม่เจอ)")
            return "Not Found"

    except Exception as e:
        print(f"❌ Error Sync Flow: {e}")
        return "Error"

# ==========================================
# Main Loop (สลับการทำงาน)
# ==========================================
if __name__ == "__main__":
    db = initialize_firestore()
    if db:
        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        try:
            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 15)
            print("🤖 บอท UAT พร้อมทำงาน! (เวอร์ชันเพิ่มระบบ Looked)")

            while True:
                # ==========================================
                # ภารกิจที่ 1: CREATE TICKET
                # ==========================================
                doc_id, doc_data = get_and_lock_ticket(db)
                if doc_id:
                    if fill_servicenow_ticket(driver, wait, doc_data):
                        mark_as_completed(db, doc_id)

                # ==========================================
                # ภารกิจที่ 2: MONITOR STATUS
                # ==========================================
                s_id, s_data = get_sync_request_task(db)
                if s_id:
                    res = check_sync_status(driver, wait, s_data)
                    
                    if res == "Resolved":
                        # ถ้าเป็น Resolved ให้ปิดงานเป็น Done
                        db.collection("incidents").document(s_id).update({
                            "status": "Completed",
                            "sync_status": "Done"
                        })
                        print(f"🎉 ตรวจพบ Resolved -> อัปเดต {s_id} เป็น Done สำเร็จ!")
                        
                    elif res != "Error":
                        # ถ้าเช็คแล้ว แต่สถานะยังไม่ใช่ Resolved (เช่น เป็น Open, Work in Progress หรือหาไม่เจอ)
                        # ให้อัปเดตเป็น "looked" เพื่อไม่ให้ดึงมาเช็คซ้ำอีก
                        db.collection("incidents").document(s_id).update({
                            "sync_status": "looked"
                        })
                        print(f"👀 สถานะคือ '{res}' -> อัปเดต {s_id} เป็น 'looked' เพื่อข้ามการเช็คซ้ำ")

                # ปรินต์จุดแสดงว่าบอทยังทำงานอยู่
                print(".", end="", flush=True)
                time.sleep(15)
                
        except Exception as e:
            print(f"❌ ไม่พบ Chrome โหมด Debug: {e}")
            input("กด Enter เพื่อปิด...")
