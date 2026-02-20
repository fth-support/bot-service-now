import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==========================================
# ส่วนที่ 1: ตั้งค่า Firebase (เตรียมพร้อมรอการเชื่อมต่อ)
# ==========================================
# หมายเหตุ: เมื่อเราได้ Database URL มาแล้ว เราจะเอาเครื่องหมาย # ออก 
# และใช้ไลบรารี pyrebase หรือ firebase-admin เพื่อดึงข้อมูลจริงครับ

# import pyrebase
# firebase_config = {
#     "apiKey": "AIzaSyCKhox_fzBzuU9n6Q4hv2fHSiXVUw7_I1U",
#     "authDomain": "service-now-79151.firebaseapp.com",
#     "projectId": "service-now-79151",
#     "databaseURL": "รอคุณเอา URL มาใส่ตรงนี้นะครับ", # <-- รออัปเดต
#     "storageBucket": "service-now-79151.firebasestorage.app",
#     "messagingSenderId": "437301948553",
#     "appId": "1:437301948553:web:3f383db66adf1fababd658"
# }
# firebase = pyrebase.initialize_app(firebase_config)
# db = firebase.database()

def get_new_ticket_from_firebase():
    """
    ฟังก์ชันจำลองการดึงข้อมูลจาก Firebase ที่มี Flag ว่าเป็นงานใหม่
    เมื่อเชื่อมต่อจริง เราจะเขียนโค้ดเช็คและดึงข้อมูลมาใส่ Dictionary แบบนี้ครับ
    """
    print("กำลังตรวจสอบข้อมูลงานใหม่จาก Firebase...")
    return {
        "caller": "MALL-0006", 
        "category": "Software",
        "impact": "5 - Minor",
        "urgency": "5 - Minor", 
        "short_desc": "ทดสอบระบบ Auto Fill จาก Bot",
        "desc": "นี่คือการทดสอบกรอกข้อมูลอัตโนมัติ\nบรรทัดที่สอง",
        "channel": "Phone",
        "state": "Acknowledged",
        "assignment_group": "FTH ISS3000 MALL",
        "reported_by": "DON-001",
        "mitigate_sla": "คุณใหญ่ 0926297894"
    }

# ==========================================
# ส่วนที่ 2: ควบคุมหน้าเว็บ ServiceNow ผ่าน Remote Debugging
# ==========================================
def fill_servicenow_ticket(data):
    print("กำลังเชื่อมต่อกับ Chrome (Port 9222)...")
    
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    
    try:
        driver = webdriver.Chrome(options=chrome_options)
        print("เชื่อมต่อสำเร็จ! กำลังเริ่มทำงานบนหน้าเว็บ...")
        wait = WebDriverWait(driver, 15)

        # --- แก้ไขส่วนนี้: สลับเข้า iframe ตั้งแต่หน้า List ---
        print("กำลังค้นหาปุ่ม New...")
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "gsft_main")))
        
        new_button = wait.until(EC.element_to_be_clickable((By.ID, "sysverb_new")))
        new_button.click()

        # พอกดปุ่ม New หน้าเว็บในกรอบจะโหลดใหม่เป็นหน้าฟอร์ม
        # เราสลับออกมาหน้าจอหลักก่อน แล้วรอสลับเข้าไปใหม่ตอนฟอร์มโหลดเสร็จเพื่อความชัวร์
        print("รอโหลดหน้าฟอร์มกรอกรายละเอียด...")
        driver.switch_to.default_content()
        time.sleep(2) # หน่วงเวลาให้เว็บเริ่มเปลี่ยนหน้า
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "gsft_main")))

        # --- เริ่มกรอกข้อมูลเหมือนเดิม ---
        print("กำลังกรอกข้อมูลตามที่ได้รับมอบหมาย...")

        # ข้อ 1: Caller
        caller_field = wait.until(EC.element_to_be_clickable((By.ID, "sys_display.incident.caller_id")))
        caller_field.clear()
        caller_field.send_keys(data["caller"])
        time.sleep(1.5)
        caller_field.send_keys(Keys.RETURN)

        # ข้อ 2: Category
        Select(driver.find_element(By.ID, "incident.category")).select_by_visible_text(data["category"])

        # ข้อ 3, 4: Impact & Urgency
        Select(driver.find_element(By.ID, "incident.impact")).select_by_visible_text(data["impact"])
        Select(driver.find_element(By.ID, "incident.urgency")).select_by_visible_text(data["urgency"])

        # ข้อ 6, 7: Short description & Description
        driver.find_element(By.ID, "incident.short_description").send_keys(data["short_desc"])
        driver.find_element(By.ID, "incident.description").send_keys(data["desc"])

        # ข้อ 8: Channel
        Select(driver.find_element(By.ID, "incident.contact_type")).select_by_visible_text(data["channel"])

        # ข้อ 9: State
        Select(driver.find_element(By.ID, "incident.state")).select_by_visible_text(data["state"])

        # ข้อ 10: Assignment group
        ag_field = driver.find_element(By.ID, "sys_display.incident.assignment_group")
        ag_field.clear()
        ag_field.send_keys(data["assignment_group"])
        time.sleep(1.5)
        ag_field.send_keys(Keys.RETURN)

        print("🎉 กรอกข้อมูลเสร็จสิ้น!")
        driver.switch_to.default_content()

    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดในการทำงาน: {e}")

# ==========================================
# Main: จุดเริ่มต้นการทำงานของโปรแกรม
# ==========================================
if __name__ == "__main__":
    # ขั้นตอนที่ 1: ไปดึงงานใหม่จาก Firebase (ตอนนี้เป็นข้อมูลจำลอง)
    ticket_data = get_new_ticket_from_firebase()
    
    # ขั้นตอนที่ 2: สั่งให้บอทเริ่มกรอกข้อมูลบน ServiceNow
    fill_servicenow_ticket(ticket_data)
    
    # เพื่อให้หน้าจอไม่ปิดไปทันทีหลังทำงานจบ
    input("กด Enter เพื่อปิดหน้าต่างนี้...")
