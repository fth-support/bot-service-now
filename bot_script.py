import time
import os
import firebase_admin
from firebase_admin import credentials, firestore
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==========================================
# ส่วนที่ 1: จัดการ Firestore (System Lock)
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
    print("\n🔍 ตรวจสอบงานใหม่ (on process)...")
    try:
        docs = db.collection("incidents") \
                 .where("status", "==", "on process") \
                 .where("is_pulled", "==", False) \
                 .limit(1).get()
        
        for doc in docs:
            doc_id = doc.id
            doc_data = doc.to_dict()
            db.collection("incidents").document(doc_id).update({"is_pulled": True})
            print(f"🔒 ล็อกงาน {doc_id} เรียบร้อย")
            return doc_id, doc_data
        return None, None
    except Exception as e:
        print(f"❌ Error Firestore: {e}")
        return None, None

def mark_as_completed(db, doc_id):
    try:
        db.collection("incidents").document(doc_id).update({"status": "Completed"})
        print(f"✅ [Firestore] อัปเดตสถานะงาน {doc_id} เป็น Completed")
    except Exception as e:
        print(f"❌ Error Firestore Update: {e}")

# ==========================================
# ส่วนที่ 2: ควบคุม ServiceNow UAT (Step-by-Step)
# ==========================================
def fill_servicenow_ticket(driver, wait, data):
    try:
        print(f"🚀 เริ่มกระบวนการสร้าง Ticket: {data.get('description', 'No Subject')}")
        
        # 1. ไปที่หน้า List URL
        list_url = "https://keristest.service-now.com/now/nav/ui/classic/params/target/incident_list.do%3Fsysparm_query%3Dactive%3Dtrue"
        driver.get(list_url)
        time.sleep(5)

        # --- ⚡ ท่าแก้ปัญหา: เคลียร์ Frame และมุดใหม่ให้ชัวร์ ⚡ ---
        driver.switch_to.default_content()
        try:
            # รอจนกว่า iframe 'gsft_main' จะโผล่มาแล้วค่อยมุด
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "gsft_main")))
            print("📥 มุดเข้า iframe (gsft_main) สำเร็จ")
        except:
            print("⚠️ ไม่เจอ iframe หรืออยู่ในหน้าตรงอยู่แล้ว")

        # --- ⚡ ท่าแก้ปัญหา: บังคับกดปุ่ม New ด้วย JavaScript ⚡ ---
        print("🖱️ กำลังบังคับคลิกปุ่ม New...")
        try:
            # ค้นหาปุ่มด้วย ID ที่ได้จาก HTML
            new_btn = wait.until(EC.presence_of_element_located((By.ID, "sysverb_new")))
            # ใช้ JavaScript คลิกเพื่อเลี่ยงปัญหาโดน element อื่นบัง
            driver.execute_script("arguments[0].click();", new_btn)
            print("✅ คลิกปุ่ม New สำเร็จ")
        except Exception as e:
            print(f"❌ หาปุ่ม New หรือคลิกไม่ได้: {e}")
            # ถ้าท่าแรกพลาด ลองกดผ่านฟังก์ชันของ ServiceNow โดยตรง
            driver.execute_script("GlideList2.get('incident').action('710fa8d1db7ff200e08070d9bf96198b', 'sysverb_new');")
            print("⚡ พยายามกดผ่าน GlideList2 Action แทน")

        # รอหน้าฟอร์มกางออก (ยังอยู่ใน iframe)
        time.sleep(5)

        # --- ส่วนการกรอกข้อมูล (เหมือนเดิม) ---
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

        # จัดการ Tab External References
        try:
            print("- สลับไป Tab External References...")
            tab_element = driver.find_element(By.XPATH, "//span[contains(text(), 'External References')]")
            driver.execute_script("arguments[0].click();", tab_element)
            time.sleep(1)

            print("- กรอก External Ref No 4...")
            ext_ref = driver.find_element(By.ID, "incident.u_extrefno4")
            ext_ref.clear()
            ext_ref.send_keys(data.get("ticket_id", ""))
        except:
            print("⚠️ มีปัญหาที่ส่วน External Reference")

        # 7. กด Save (Submit)
        print("💾 กำลังกดบันทึก (Submit)...")
        driver.find_element(By.ID, "sysverb_insert").click()
        
        time.sleep(5) 
        return True 

    except Exception as e:
        print(f"❌ Error ServiceNow Flow: {e}")
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
            print("🤖 บอทพร้อมรันระบบ UAT! (โหมด Iframe Handling)")

            while True:
                doc_id, doc_data = get_and_lock_ticket(db)
                if doc_id:
                    if fill_servicenow_ticket(driver, wait, doc_data):
                        mark_as_completed(db, doc_id)
                else:
                    print(".", end="", flush=True)
                time.sleep(15)
        except Exception as e:
            print(f"❌ ไม่พบ Chrome โหมด Debug: {e}")
            input("กด Enter เพื่อปิด...")
