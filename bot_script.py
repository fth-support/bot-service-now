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
# ส่วนที่ 1: จัดการ Firestore (ปรับปรุงระบบ Lock)
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

def get_and_lock_ticket(db):
    """ ค้นหาข้อมูลและทำการ 'Lock' ทันทีเพื่อป้องกันการทำงานซ้ำ """
    print("\n🔍 กำลังตรวจสอบงานใหม่...")
    try:
        # 1. ค้นหางานที่ยังไม่ถูกดึง (is_pulled == False)
        docs = db.collection("incidents") \
                 .where("status", "==", "Pending") \
                 .where("is_pulled", "==", False) \
                 .limit(1).get()
        
        for doc in docs:
            doc_id = doc.id
            doc_data = doc.to_dict()
            
            # 2. ทำการ STAMP ทันที! ว่างานนี้ถูกดึงไปแล้ว (Locking Mechanism)
            # แม้บอทจะแครชหลังจากนี้ งานนี้จะไม่ถูกหยิบมาทำซ้ำ
            db.collection("incidents").document(doc_id).update({"is_pulled": True})
            print(f"🔒 ล็อกงาน {doc_id} เรียบร้อย (Status: In-Progress)")
            return doc_id, doc_data
            
        return None, None
    except Exception as e:
        print(f"❌ Error Firestore (Read/Lock): {e}")
        return None, None

def mark_as_completed(db, doc_id):
    """ อัปเดตสถานะสุดท้ายเมื่อบันทึกบน ServiceNow สำเร็จ """
    try:
        db.collection("incidents").document(doc_id).update({"status": "Completed"})
        print(f"✅ [Firestore] ปิดงาน {doc_id} สมบูรณ์")
    except Exception as e:
        print(f"❌ Error Firestore (Complete): {e}")

# ==========================================
# ส่วนที่ 2: ควบคุม ServiceNow (ปรับปรุงการ Force Select)
# ==========================================
def fill_servicenow_ticket(driver, wait, data):
    try:
        print(f"🚀 เริ่มสร้าง Ticket: {data.get('short_description', 'No Subject')}")
        driver.get("https://keris.service-now.com/incident.do?sys_id=-1")
        time.sleep(4) 

        # --- ข้อ 1: Force Select 'Sense & Respond Demand' ---
        print("- กำลังเลือก Sense & Respond Demand...")
        try:
            # ใช้ JavaScript เพื่อบังคับเลือกค่า แม้หน้าเว็บจะ Responsive
            js_script = """
            var select = document.getElementById('incident.u_sense_respond_demand');
            for (var i = 0; i < select.options.length; i++) {
                if (select.options[i].text === 'Retail Solution & Delivery') {
                    select.selectedIndex = i;
                    select.dispatchEvent(new Event('change'));
                    break;
                }
            }
            """
            driver.execute_script(js_script)
            time.sleep(1)
        except Exception as ex:
            print(f"  ⚠️ เตือน: ไม่สามารถเลือก Demand ได้อัตโนมัติ: {ex}")

        # กรอก Caller
        caller_field = wait.until(EC.element_to_be_clickable((By.ID, "sys_display.incident.caller_id")))
        caller_field.clear()
        caller_field.send_keys(data.get("caller", ""))
        time.sleep(1.5)
        caller_field.send_keys(Keys.RETURN)
        
        # กรอกฟิลด์หลัก
        driver.find_element(By.ID, "incident.short_description").send_keys(data.get("short_description", ""))
        driver.find_element(By.ID, "incident.description").send_keys(data.get("description", ""))
        
        # เลือก Category, Impact, Urgency
        Select(driver.find_element(By.ID, "incident.category")).select_by_visible_text(data.get("category", "Software"))
        Select(driver.find_element(By.ID, "incident.impact")).select_by_visible_text(data.get("impact", "5 - Minor"))
        Select(driver.find_element(By.ID, "incident.urgency")).select_by_visible_text(data.get("urgency", "5 - Minor"))
        
        # Assignment Group
        ag_field = driver.find_element(By.ID, "sys_display.incident.assignment_group")
        ag_field.clear()
        ag_field.send_keys("FTH ISS3000 MALL")
        time.sleep(1.5)
        ag_field.send_keys(Keys.RETURN)

        # Mitigate SLA (contact_info)
        driver.find_element(By.ID, "incident.u_mitigate_sla_description").send_keys(data.get("contact_info", ""))

        # กด Save
        print("💾 กำลังกดบันทึก...")
        driver.find_element(By.ID, "sysverb_insert").click()
        
        time.sleep(5) 
        return True 

    except Exception as e:
        print(f"❌ Error ServiceNow: {e}")
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
            print("🤖 บอทพร้อมทำงาน (โหมดป้องกันการทำซ้ำ)")

            while True:
                # 1. ค้นหางานและ 'ล็อก' ทันที
                doc_id, doc_data = get_and_lock_ticket(db)
                
                if doc_id:
                    # 2. ทำงานบนหน้าเว็บ
                    if fill_servicenow_ticket(driver, wait, doc_data):
                        # 3. ถ้าสำเร็จ ค่อยเปลี่ยน Status เป็น Completed
                        mark_as_completed(db, doc_id)
                    else:
                        print(f"⚠️ งาน {doc_id} ทำงานไม่สำเร็จบนเว็บ แต่ถูกล็อกไว้แล้วเพื่อความปลอดภัย")
                else:
                    print(".", end="", flush=True)
                
                time.sleep(15)

        except Exception as e:
            print(f"❌ ไม่สามารถเชื่อมต่อกับ Chrome ได้: {e}")
            input("กด Enter เพื่อปิด...")
