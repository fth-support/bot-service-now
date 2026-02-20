import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==========================================
# ส่วนที่ 1: จัดการข้อมูล Firebase
# ==========================================
def get_new_ticket_from_firebase():
    print("กำลังตรวจสอบข้อมูลงานใหม่จาก Firebase...")
    # สมมติว่านี่คือข้อมูลที่ดึงมาจาก Firebase ของจริง
    # เราจะต้องส่ง "รหัสอ้างอิง (ticket_id)" กลับมาด้วยเสมอ
    mock_ticket_id = "RECORD_ID_001" 
    mock_data = {
        "caller": "MALL-0006", 
        "category": "Software",
        "impact": "5 - Minor",
        "urgency": "5 - Minor", 
        "short_desc": "ทดสอบระบบ Auto Fill จาก Bot พร้อมส่ง Flag กลับ",
        "desc": "นี่คือการทดสอบกรอกข้อมูลอัตโนมัติ",
        "channel": "Phone",
        "state": "Acknowledged",
        "assignment_group": "FTH ISS3000 MALL"
    }
    return mock_ticket_id, mock_data

def update_flag_in_firebase(ticket_id):
    """
    ฟังก์ชันสำหรับส่ง Flag กลับไปบอก Firebase ว่าเปิดตั๋วสำเร็จแล้ว
    """
    print(f"กำลังส่ง Flag กลับไปที่ Firebase สำหรับงาน: {ticket_id}...")
    
    try:
        # โค้ดของจริงจะหน้าตาประมาณนี้ครับ (รอเอาเครื่องหมาย # ออกเมื่อเชื่อมต่อจริง)
        # db.child("YOUR_COLLECTION_NAME").child(ticket_id).update({"status": "completed"})
        
        print("✅ อัปเดตสถานะบน Firebase เป็น 'completed' สำเร็จ!")
    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดในการอัปเดต Firebase: {e}")

# ==========================================
# ส่วนที่ 2: ควบคุมหน้าเว็บ ServiceNow 
# ==========================================
def fill_servicenow_ticket(data):
    print("กำลังเชื่อมต่อกับ Chrome (Port 9222)...")
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    
    try:
        driver = webdriver.Chrome(options=chrome_options)
        wait = WebDriverWait(driver, 15)

        print("กำลังค้นหาและคลิกปุ่ม New...")
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "gsft_main")))
        new_button = wait.until(EC.element_to_be_clickable((By.ID, "sysverb_new")))
        new_button.click()

        print("รอโหลดหน้าฟอร์มกรอกรายละเอียด...")
        driver.switch_to.default_content()
        time.sleep(2) 
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "gsft_main")))

        print("กำลังกรอกข้อมูล...")
        
        # ตัวอย่างการกรอกข้อมูล 1 ช่องเพื่อความรวดเร็วในการแสดงตัวอย่าง
        caller_field = wait.until(EC.element_to_be_clickable((By.ID, "sys_display.incident.caller_id")))
        caller_field.clear()
        caller_field.send_keys(data["caller"])
        time.sleep(1.5)
        caller_field.send_keys(Keys.RETURN)
        
        # ... (โค้ดกรอกข้อมูลฟิลด์อื่นๆ ใส่ไว้ตรงนี้เหมือนเดิมครับ) ...

        print("🎉 บอททำงานบน ServiceNow เสร็จสิ้น!")
        driver.switch_to.default_content()
        
        # สำคัญมาก: ต้องส่งค่า True กลับไปเพื่อบอกว่ากรอกสำเร็จและไม่พังกลางทาง
        return True 

    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดบนหน้าเว็บ: {e}")
        # ถ้ามี Error จะส่งค่า False กลับไป เพื่อไม่ให้อัปเดต Firebase มั่วๆ
        return False

# ==========================================
# Main: จุดเริ่มต้นการทำงานของโปรแกรม
# ==========================================
if __name__ == "__main__":
    # 1. ไปดึงงานใหม่ และ ID ของงานนั้นมาจาก Firebase
    ticket_id, ticket_data = get_new_ticket_from_firebase()
    
    # 2. สั่งให้บอทเริ่มกรอกข้อมูลบน ServiceNow
    is_success = fill_servicenow_ticket(ticket_data)
    
    # 3. ตรวจสอบว่าถ้าเปิดตั๋วสำเร็จ ค่อยส่ง Flag ไปอัปเดต Firebase
    if is_success:
        update_flag_in_firebase(ticket_id)
    else:
        print("⚠️ ข้ามการอัปเดต Firebase เนื่องจากบอททำงานบนหน้าเว็บไม่สำเร็จ")
    
    input("กด Enter เพื่อปิดหน้าต่างนี้...")
