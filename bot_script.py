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
# ส่วนที่ 1: จัดการ Firestore (Locking System)
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
    print("\n🔍 กำลังตรวจสอบงานใหม่...")
    try:
        docs = db.collection("incidents") \
                 .where("status", "==", "Pending") \
                 .where("is_pulled", "==", False) \
                 .limit(1).get()
        
        for doc in docs:
            doc_id = doc.id
            doc_data = doc.to_dict()
            db.collection("incidents").document(doc_id).update({"is_pulled": True})
            print(f"🔒 ล็อกงาน {doc_id} เรียบร้อย (Status: In-Progress)")
            return doc_id, doc_data
        return None, None
    except Exception as e:
        print(f"❌ Error Firestore (Read/Lock): {e}")
        return None, None

def mark_as_completed(db, doc_id):
    try:
        db.collection("incidents").document(doc_id).update({"status": "Completed"})
        print(f"✅ [Firestore] ปิดงาน {doc_id} สมบูรณ์")
    except Exception as e:
        print(f"❌ Error Firestore (Complete): {e}")

# ==========================================
# ส่วนที่ 2: ควบคุม ServiceNow
# ==========================================
def fill_servicenow_ticket(driver, wait, data):
    try:
        print(f"🚀 เริ่มสร้าง Ticket: {data.get('short_description', 'No Subject')}")
        driver.get("https://keris.service-now.com/incident.do?sys_id=-1")
        
        # รอให้ฟิลด์แรกปรากฏขึ้นมาเพื่อยืนยันว่าหน้าโหลดเสร็จ
        wait.until(EC.presence_of_element_located((By.ID, "incident.short_description")))

        # --- แก้ไขจุดที่ 1: ใช้ท่า Select ปกติ (เหมือน Impact/Urgency) ---
        print("- กำลังเลือก Sense & Respond Demand...")
        try:
            # ใช้ WebDriverWait รอให้ Element พร้อมที่จะเลือก
            demand_element = wait.until(EC.element_to_be_clickable((By.ID, "incident.u_sense_respond_demand")))
            dropdown = Select(demand_element)
            dropdown.select_by_visible_text("Retail Solution & Delivery")
            print("  ✅ เลือก Demand สำเร็จ")
        except Exception as ex:
            print(f"  ⚠️ ไม่สามารถเลือก Demand ได้ (อาจจะ ID ไม่ตรงหรือตัวเลือกไม่มี): {ex}")

        # พิมพ์ Caller
        caller_field = wait.until(EC.element_to_be_clickable((By.ID, "sys_display.incident.caller_id")))
        caller_field.clear()
        caller_field.send_keys(data.get("caller", ""))
        time.sleep(2)
        caller_field.send_keys(Keys.RETURN)
        
        # กรอกฟิลด์ข้อความ
        driver.find_element(By.ID, "incident.short_description").send_keys(data.get("short_description", ""))
        driver.find_element(By.ID, "incident.description").send_keys(data.get("description", ""))
        
        # เลือก Category, Impact, Urgency (ใช้ท่ามาตรฐานที่ได้ผลเสมอ)
        try:
            Select(driver.find_element(By.ID, "incident.category")).select_by_visible_text(data.get("category", "Software"))
            Select(driver.find_element(By.ID, "incident.impact")).select_by_visible_text(data.get("impact", "5 - Minor"))
            Select(driver.find_element(By.ID, "incident.urgency")).select_by_visible_text(data.get("urgency", "5 - Minor"))
            print("  ✅ เลือก Category/Impact/Urgency สำเร็จ")
        except Exception as e:
            print(f"  ⚠️ Error ในการเลือก Dropdown มาตรฐาน: {e}")
        
        # Assignment Group
        ag_field = driver.find_element(By.ID, "sys_display.incident.assignment_group")
        ag_field.clear()
        ag_field.send_keys("FTH ISS3000 MALL")
        time.sleep(2)
        ag_field.send_keys(Keys.RETURN)

        # --- จุดที่ 2: Reported By (Reference Field ที่แก้แล้วได้ผล) ---
        print("- กำลังกรอก Reported By...")
        try:
            rep_field = wait.until(EC.element_to_be_clickable((By.ID, "sys_display.incident.u_reported_by")))
            rep_field.clear()
            rep_field.send_keys(data.get("reported_by", "DON-001"))
            time.sleep(2)
            rep_field.send_keys(Keys.RETURN)
            print("  ✅ กรอก Reported By สำเร็จ")
        except Exception as ex:
            print(f"  ⚠️ ไม่สามารถกรอก Reported By ได้: {ex}")

        # Mitigate SLA (contact_info)
        driver.find_element(By.ID, "incident.u_mitigate_sla_description").send_keys(data.get("contact_info", ""))

        # กด Save
        print("💾 กำลังกดบันทึก...")
        driver.find_element(By.ID, "sysverb_insert").click()
        
        time.sleep(5) 
        return True 

    except Exception as e:
        print(f"❌ Error ServiceNow Main Flow: {e}")
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
            print("🤖 บอทเฝ้าระบบพร้อมทำงาน...")

            while True:
                doc_id, doc_data = get_and_lock_ticket(db)
                if doc_id:
                    if fill_servicenow_ticket(driver, wait, doc_data):
                        mark_as_completed(db, doc_id)
                    else:
                        print(f"⚠️ งาน {doc_id} มีปัญหาบนหน้าเว็บ (ตรวจสอบสถานะ is_pulled ใน Firestore)")
                else:
                    print(".", end="", flush=True)
                time.sleep(15)
        except Exception as e:
            print(f"❌ ไม่สามารถเชื่อมต่อกับ Chrome ได้: {e}")
            input("กด Enter เพื่อปิด...")
