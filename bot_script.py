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
# ส่วนที่ 1: จัดการ Firestore (Rollback กลับมาใช้แบบเดิมที่เสถียร)
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

# --- ดึงงาน Create ---
def get_and_lock_ticket(db):
    try:
        # ใช้รูปแบบเดิมของคุณที่ทำงานได้ชัวร์ 100%
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

# --- ดึงงาน Sync/Track ---
def get_sync_request_task(db):
    try:
        # ใช้รูปแบบเดิมของคุณ
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
# ส่วนที่ 2: ควบคุม ServiceNow UAT (แยกงานชัดเจน)
# ==========================================

# --- งานที่ 2.1: สร้าง Ticket (โค้ดเดิมของคุณ 100%) ---
def fill_servicenow_ticket(driver, wait, data):
    try:
        print(f"🚀 เริ่มสร้าง Ticket: {data.get('description', 'No Subject')}")
        
        # บังคับหน้าจอให้เป็นปัจจุบัน
        driver.switch_to.default_content()
        new_incident_url = "https://keristest.service-now.com/incident.do?sys_id=-1"
        driver.get(new_incident_url)
        
        print("⏳ รอหน้าฟอร์มกางตัว...")
        time.sleep(6) 

        # ตรวจสอบการมุด iframe 'gsft_main'
        try:
            if not driver.find_elements(By.ID, "incident.short_description"):
                wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "gsft_main")))
                print("📥 มุดเข้า iframe (gsft_main) สำเร็จ")
        except:
            pass

        # 1. Short Description
        print("- กรอก Short Description...")
        short_desc = wait.until(EC.presence_of_element_located((By.ID, "incident.short_description")))
        short_desc.clear()
        short_desc.send_keys(data.get("description", ""))

        # 2. Caller
        print("- กรอก Caller...")
        caller_field = driver.find_element(By.ID, "sys_display.incident.caller_id")
        caller_field.clear()
        caller_field.send_keys(data.get("caller", ""), Keys.RETURN)
        time.sleep(2)

        # 3. Assignment Group
        print("- กรอก Assignment Group...")
        ag_field = driver.find_element(By.ID, "sys_display.incident.assignment_group")
        ag_field.clear()
        ag_field.send_keys("FTH Call Center", Keys.RETURN)
        time.sleep(2)

        # 4. Category, Impact, Urgency
        print("- กำลังตั้งค่า Category, Impact, Urgency...")
        try:
            cat_value = data.get("category", "Software") 
            Select(driver.find_element(By.ID, "incident.category")).select_by_visible_text(cat_value)
            Select(driver.find_element(By.ID, "incident.impact")).select_by_visible_text("5 - Minor")
            Select(driver.find_element(By.ID, "incident.urgency")).select_by_visible_text("5 - Minor")
            print("  ✅ ตั้งค่า Dropdowns สำเร็จ")
        except Exception as e:
            print(f"  ⚠️ ไม่สามารถเลือก Dropdown ได้: {e}")

        # 5. สลับไป Tab External Reference และกรอกข้อมูล
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

        # 6. กด Save
        print("💾 กำลังกดบันทึก...")
        driver.find_element(By.ID, "sysverb_insert").click()
        
        time.sleep(5) 
        return True 

    except Exception as e:
        print(f"❌ Error ServiceNow Flow: {e}")
        return False

# --- งานที่ 2.2: Track / Monitor Status (ใช้ URL สั้น) ---
def check_sync_status(driver, wait, data):
    try:
        ticket_id = data.get("ticket_id")
        print(f"🔍 [Monitor] ค้นหาสถานะของ: {ticket_id}")
        
        # ใช้ URL สั้นตามคำแนะนำ เพื่อตัดปัญหาต่อท้ายที่อาจเปลี่ยนไป
        driver.switch_to.default_content()
        driver.get("https://keristest.service-now.com/incident_list.do")
        time.sleep(5)

        # เช็คว่ามี Iframe ให้มุดหรือไม่ (บางที URL สั้นอาจไม่มี Iframe)
        try:
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "gsft_main")))
        except:
            pass

        # 1. เลือก Dropdown "for text" 
        # ใช้ XPATH หากล่องที่ครอบ Dropdown ไว้แบบไม่ต้องสน ID
        print("- เลือกโหมด 'for text'...")
        search_dropdown_element = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(@class, 'input-group-select')]//select")))
        Select(search_dropdown_element).select_by_value("zztextsearchyy")
        time.sleep(1)

        # 2. กรอกเลข INC และกด Enter
        # หาช่อง Input ที่อยู่ติดกัน
        print(f"- กรอกเลขค้นหา: {ticket_id}")
        search_input = driver.find_element(By.XPATH, "//div[contains(@class, 'input-group')]//input[@type='search']")
        search_input.clear()
        search_input.send_keys(ticket_id)
        time.sleep(1)
        search_input.send_keys(Keys.ENTER)
        
        print("⏳ รอระบบโหลดผลลัพธ์...")
        time.sleep(5) 

        # 3. เช็คสถานะ State ในแถวแรก
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
            print("🤖 บอท UAT พร้อมทำงาน! (เวอร์ชัน Stable แยก Create กับ Monitor)")

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
                        db.collection("incidents").document(s_id).update({
                            "status": "Completed",
                            "sync_status": "Done"
                        })
                        print(f"🎉 ตรวจพบ Resolved -> อัปเดต {s_id} สำเร็จ!")
                    else:
                        print(f"⏳ {s_id} ยังเป็นสถานะ '{res}' (ข้ามไปก่อน)")

                # ปรินต์จุดแสดงว่าบอทยังทำงานอยู่
                print(".", end="", flush=True)
                time.sleep(15)
                
        except Exception as e:
            print(f"❌ ไม่พบ Chrome โหมด Debug: {e}")
            input("กด Enter เพื่อปิด...")
