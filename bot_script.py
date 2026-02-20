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

        # ---------------------------------------------------------
        # 🚀 ท่าไม้ตายขั้นสุดยอด: ยิงเข้าหน้า Raw Form โดยตรง (ไม่มี iframe)
        # ---------------------------------------------------------
        print("กำลังพาไปหน้า Raw Form (ลดการประมวลผล iframe)...")
        # สังเกตว่า URL จะสั้นลง เป็นการเรียกฟอร์มแบบตรงๆ
        raw_form_url = "https://keris.service-now.com/incident.do?sys_id=-1"
        driver.get(raw_form_url)

        print("รอโหลดหน้าฟอร์มกรอกรายละเอียด...")
        time.sleep(3) # หน่วงเวลาให้เว็บโหลดเสร็จ
        
        # ❌ เราไม่ต้องใช้คำสั่ง switch_to.frame อีกต่อไปแล้ว เพราะหน้านี้ไม่มีกรอบ!
        
        print("กำลังเริ่มกรอกข้อมูล...")
        
        # --- (เริ่มกรอกข้อมูลทันที) ---
        caller_field = wait.until(EC.element_to_be_clickable((By.ID, "sys_display.incident.caller_id")))
        caller_field.clear()
        caller_field.send_keys(data["caller"])
        time.sleep(1.5)
        caller_field.send_keys(Keys.RETURN)
        
        Select(driver.find_element(By.ID, "incident.category")).select_by_visible_text(data["category"])
        Select(driver.find_element(By.ID, "incident.impact")).select_by_visible_text(data["impact"])
        Select(driver.find_element(By.ID, "incident.urgency")).select_by_visible_text(data["urgency"])
        driver.find_element(By.ID, "incident.short_description").send_keys(data["short_desc"])
        driver.find_element(By.ID, "incident.description").send_keys(data["desc"])
        Select(driver.find_element(By.ID, "incident.contact_type")).select_by_visible_text(data["channel"])
        Select(driver.find_element(By.ID, "incident.state")).select_by_visible_text(data["state"])
        
        ag_field = driver.find_element(By.ID, "sys_display.incident.assignment_group")
        ag_field.clear()
        ag_field.send_keys(data["assignment_group"])
        time.sleep(1.5)
        ag_field.send_keys(Keys.RETURN)

        print("🎉 บอทกรอกข้อมูลบน ServiceNow เสร็จสิ้น!")
        # ❌ ไม่ต้อง switch_to.default_content() แล้วเช่นกัน
        
        return True 

    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดบนหน้าเว็บ: {e}")
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
