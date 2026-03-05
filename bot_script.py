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
        docs = db.collection("incidents") \
                 .where("sync_status", "==", "Request") \
                 .limit(1).get()
                 
        for doc in docs:
            return doc.id, doc.to_dict()
        return None, None
    except Exception as e:
        print(f"❌ Error Firestore Sync: {e}")
        return None, None

# ==========================================
# ส่วนที่ 2: ควบคุม ServiceNow UAT
# ==========================================

# ท่าบังคับ: ดึงบอทกลับมา Tab ปัจจุบัน
def force_active_tab(driver):
    try:
        for handle in driver.window_handles:
            driver.switch_to.window(handle)
            if driver.current_url.startswith("http"):
                driver.maximize_window()
                return
                
        driver.execute_script("window.open('https://keristest.service-now.com', '_blank');")
        driver.switch_to.window(driver.window_handles[-1])
        driver.maximize_window()
    except Exception as e:
        print(f"⚠️ Error forcing tab: {e}")

# --- งานที่ 2.1: สร้าง Ticket ---
def fill_servicenow_ticket(driver, wait, data):
    try:
        print(f"🚀 เริ่มสร้าง Ticket: {data.get('description', 'No Subject')}")
        
        force_active_tab(driver)
        driver.get("https://keristest.service-now.com/incident.do?sys_id=-1")
        
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
        
        # ยิง URL ค้นหาโดยตรง
        search_url = f"https://keristest.service-now.com/incident_list.do?sysparm_query=GOTO123TEXTQUERY321%3d{ticket_id}"
        print(f"- ยิง URL ค้นหาโดยตรง...")
        driver.get(search_url)
        time.sleep(6) # รอหน้าเว็บโหลดผลลัพธ์

        driver.switch_to.default_content()
        if driver.find_elements(By.ID, "gsft_main"):
            driver.switch_to.frame("gsft_main")

        # ⚡ ปรับแก้: หาคำว่า "Resolved" แบบกวาดสายตาตามที่คุณแนะนำ
        try:
            # ดึงข้อความทั้งหมดบนหน้าเว็บ (หรือในตารางผลลัพธ์) มาเช็คเลยว่ามีคำว่า Resolved ไหม
            page_text = driver.find_element(By.TAG_NAME, "body").text
            
            if "Resolved" in page_text:
                print("📍 พบคำว่า 'Resolved' ในหน้าเว็บ!")
                return "Resolved"
            else:
                print("📍 ไม่พบคำว่า 'Resolved' (สถานะอาจยังเป็น Open)")
                return "Open"
        except:
            print("⚠️ ไม่สามารถอ่านข้อมูลหน้าเว็บได้")
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
            print("🤖 บอท UAT พร้อมทำงาน! (อัปเกรดระบบตรวจจับ Resolved)")

            while True:
                # ==========================================
                # ภารกิจที่ 1: CREATE TICKET
                # ==========================================
                doc_id, doc_data = get_and_lock_ticket(db)
                if doc_id:
                    if fill_servicenow_ticket(driver, wait, doc_data):
                        print(f"✅ คีย์ Ticket เข้าเว็บเรียบร้อย (คงสถานะ 'on process' ไว้)")

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
                        print(f"🎉 ตรวจพบ Resolved -> อัปเดต {s_id} เป็น Done และ Completed สำเร็จ!")
                        
                    elif res != "Error":
                        db.collection("incidents").document(s_id).update({
                            "sync_status": "looked"
                        })
                        print(f"👀 ไม่ใช่ Resolved -> อัปเดต {s_id} เป็น 'looked' เพื่อข้ามการเช็คซ้ำ")

                # ปรินต์จุดแสดงว่าบอทยังทำงานอยู่
                print(".", end="", flush=True)
                time.sleep(15)
                
        except Exception as e:
            print(f"❌ ไม่พบ Chrome โหมด Debug: {e}")
            input("กด Enter เพื่อปิด...")
