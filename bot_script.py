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
# ส่วนที่ 1: ตั้งค่า Firestore (กุญแจสำคัญ)
# ==========================================
def initialize_firestore():
    try:
        # ตรวจสอบว่ามีไฟล์กุญแจ serviceAccount.json อยู่ในโฟลเดอร์เดียวกับ .exe หรือไม่
        if not os.path.exists("serviceAccount.json"):
            print("❌ ไม่พบไฟล์ serviceAccount.json กรุณานำไฟล์กุญแจมาวางไว้ที่เดียวกับบอท")
            return None
        
        cred = credentials.Certificate("serviceAccount.json")
        firebase_admin.initialize_app(cred)
        return firestore.client()
    except Exception as e:
        print(f"❌ ไม่สามารถเชื่อมต่อ Firestore ได้: {e}")
        return None

def get_new_ticket_from_firestore(db):
    print("\n🔍 กำลังตรวจสอบงานใหม่จาก Firestore (incidents)...")
    try:
        # ค้นหาใน collection 'incidents' ที่ status == 'Pending' และ is_pulled == False
        docs = db.collection("incidents") \
                 .where("status", "==", "Pending") \
                 .where("is_pulled", "==", False) \
                 .limit(1).get()
        
        for doc in docs:
            return doc.id, doc.to_dict()
        return None, None
    except Exception as e:
        print(f"❌ Error Firestore (Read): {e}")
        return None, None

def update_flag_in_firestore(db, doc_id):
    try:
        # อัปเดตสถานะเป็น Completed และป้องกันการดึงซ้ำด้วย is_pulled = True
        db.collection("incidents").document(doc_id).update({
            "status": "Completed",
            "is_pulled": True
        })
        print(f"✅ [Firestore] อัปเดตงาน {doc_id} เป็น Completed เรียบร้อย!")
    except Exception as e:
        print(f"❌ Error Firestore (Update): {e}")

# ==========================================
# ส่วนที่ 2: ควบคุม ServiceNow (Raw Form)
# ==========================================
def fill_servicenow_ticket(driver, wait, data):
    try:
        print(f"🚀 เริ่มสร้าง Ticket: {data.get('short_description', 'No Subject')}")
        driver.get("https://keris.service-now.com/incident.do?sys_id=-1")
        time.sleep(4) 

        # 1. Sense & Respond Demand
        try:
            demand_select = Select(wait.until(EC.presence_of_element_located((By.ID, "incident.u_sense_respond_demand"))))
            demand_select.select_by_visible_text("Retail Solution & Delivery")
        except: pass

        # 2. Caller (ดึงจาก Firebase)
        caller_field = wait.until(EC.element_to_be_clickable((By.ID, "sys_display.incident.caller_id")))
        caller_field.clear()
        caller_field.send_keys(data.get("caller", ""))
        time.sleep(1.5)
        caller_field.send_keys(Keys.RETURN)
        
        # 3. Short Description & Description
        driver.find_element(By.ID, "incident.short_description").send_keys(data.get("short_description", ""))
        driver.find_element(By.ID, "incident.description").send_keys(data.get("description", ""))
        
        # 4. Dropdowns ต่างๆ
        Select(driver.find_element(By.ID, "incident.category")).select_by_visible_text(data.get("category", "Software"))
        Select(driver.find_element(By.ID, "incident.impact")).select_by_visible_text(data.get("impact", "5 - Minor"))
        Select(driver.find_element(By.ID, "incident.urgency")).select_by_visible_text(data.get("urgency", "5 - Minor"))
        
        # 5. Assignment Group
        ag_field = driver.find_element(By.ID, "sys_display.incident.assignment_group")
        ag_field.clear()
        ag_field.send_keys("FTH ISS3000 MALL") # หรือดึงจาก data ถ้ามี
        time.sleep(1.5)
        ag_field.send_keys(Keys.RETURN)

        # 6. Mitigate SLA Description (ใช้ contact_info)
        driver.find_element(By.ID, "incident.u_mitigate_sla_description").send_keys(data.get("contact_info", ""))

        # 7. กด Save (Submit)
        print("💾 กำลังกดบันทึก...")
        driver.find_element(By.ID, "sysverb_insert").click()
        
        time.sleep(5) # รอหน้า All Job โหลด
        return True 

    except Exception as e:
        print(f"❌ Error ServiceNow: {e}")
        return False

# ==========================================
# Main: ระบบ Loop
# ==========================================
if __name__ == "__main__":
    db = initialize_firestore()
    
    if db:
        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        
        try:
            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 15)
            print("🤖 บอทพร้อมทำงาน! กำลังเฝ้าดูงานใหม่ใน Firestore...")

            while True:
                doc_id, doc_data = get_new_ticket_from_firestore(db)
                
                if doc_id:
                    if fill_servicenow_ticket(driver, wait, doc_data):
                        update_flag_in_firestore(db, doc_id)
                else:
                    print(".", end="", flush=True)
                
                time.sleep(15) # พัก 15 วินาทีก่อนเช็คใหม่

        except Exception as e:
            print(f"❌ ไม่สามารถเชื่อมต่อกับ Chrome ได้: {e}")
            input("กด Enter เพื่อปิด...")
