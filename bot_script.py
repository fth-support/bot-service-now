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
    # นี่คือรูปแบบข้อมูลที่คุณแจ้งมาล่าสุด
    mock_ticket_id = "FB_RECORD_12345" 
    mock_data = {
        "caller": "MALL-0006", 
        "category": "Software",
        "impact": "5 - Minor",
        "urgency": "5 - Minor", 
        "short_description": "ดึงข้อมูลจาก firebase: ปัญหาระบบ POS", # แก้ไขตามข้อ 1
        "description": "ดึงข้อมูลจาก firebase: รายละเอียดปัญหาอย่างละเอียด...", # แก้ไขตามข้อ 2
        "contact_info": "คุณใหญ่ 0926297894", # แก้ไขตามข้อ 3
        "channel": "Phone",
        "state": "Acknowledged",
        "assignment_group": "FTH ISS3000 MALL",
        "reported_by": "DON-001"
    }
    return mock_ticket_id, mock_data

def update_flag_in_firebase(ticket_id):
    """ ส่งสถานะกลับไปบอก Firebase ว่าเปิด ticket สำเร็จ (แก้ไขตามข้อ 4 ส่วนที่ 3) """
    print(f"✅ ส่งสถานะ: งาน {ticket_id} เปิด Ticket บน ServiceNow สำเร็จแล้ว!")
    # โค้ดจริง: db.child("tickets").child(ticket_id).update({"status": "completed"})

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

        print("กำลังพาไปหน้า Raw Form...")
        raw_form_url = "https://keris.service-now.com/incident.do?sys_id=-1"
        driver.get(raw_form_url)
        time.sleep(3) 
        
        print("กำลังเริ่มกรอกข้อมูล...")

        # --- แก้ไขข้อ 4: Auto Select 'Sense & Respond Demand' ---
        try:
            # หมายเหตุ: ID ช่องนี้มักจะเป็น incident.u_sense_respond_demand (ตรวจสอบจากหน้าจออีกที)
            demand_select = Select(wait.until(EC.presence_of_element_located((By.ID, "incident.u_sense_respond_demand"))))
            demand_select.select_by_visible_text("Retail Solution & Delivery")
            print("- เลือก Sense & Respond Demand สำเร็จ")
        except:
            print("- ไม่พบฟิลด์ Sense & Respond Demand หรือ ID ไม่ตรง")

        # ข้อ 1: Caller
        caller_field = wait.until(EC.element_to_be_clickable((By.ID, "sys_display.incident.caller_id")))
        caller_field.clear()
        caller_field.send_keys(data["caller"])
        time.sleep(1.5)
        caller_field.send_keys(Keys.RETURN)
        
        # ฟิลด์มาตรฐานอื่นๆ
        Select(driver.find_element(By.ID, "incident.category")).select_by_visible_text(data["category"])
        Select(driver.find_element(By.ID, "incident.impact")).select_by_visible_text(data["impact"])
        Select(driver.find_element(By.ID, "incident.urgency")).select_by_visible_text(data["urgency"])

        # --- แก้ไขข้อ 1 & 2: ดึงข้อมูลให้ถูก Key จาก Firebase ---
        driver.find_element(By.ID, "incident.short_description").send_keys(data["short_description"])
        driver.find_element(By.ID, "incident.description").send_keys(data["description"])

        Select(driver.find_element(By.ID, "incident.contact_type")).select_by_visible_text(data["channel"])
        Select(driver.find_element(By.ID, "incident.state")).select_by_visible_text(data["state"])
        
        # Assignment Group & Reported By
        ag_field = driver.find_element(By.ID, "sys_display.incident.assignment_group")
        ag_field.clear()
        ag_field.send_keys(data["assignment_group"])
        time.sleep(1.5)
        ag_field.send_keys(Keys.RETURN)

        rep_field = driver.find_element(By.ID, "sys_display.incident.u_reported_by")
        rep_field.clear()
        rep_field.send_keys(data["reported_by"])
        time.sleep(1.5)
        rep_field.send_keys(Keys.RETURN)

        # --- แก้ไขข้อ 3: Mitigate SLA Description (ใช้ contact_info จาก firebase) ---
        # ID มักจะเป็น incident.u_mitigate_sla_description
        driver.find_element(By.ID, "incident.u_mitigate_sla_description").send_keys(data["contact_info"])

        # --- แก้ไขข้อ 5: กดปุ่ม Save เพื่อยืนยันและกลับหน้าหลัก ---
        print("กำลังกดบันทึก (Save)...")
        # ปุ่ม Save ใน SNOW ปกติ ID คือ sysverb_insert
        save_button = driver.find_element(By.ID, "sysverb_insert")
        save_button.click()
        
        print("🎉 บอทบันทึกงานสำเร็จและกำลังกลับหน้า All Job!")
        return True 

    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาด: {e}")
        return False

# ==========================================
# Main
# ==========================================
if __name__ == "__main__":
    # 1. ดึงงาน
    ticket_id, ticket_data = get_new_ticket_from_firebase()
    
    # 2. กรอกและบันทึก
    if fill_servicenow_ticket(ticket_data):
        # 3. ถ้าสำเร็จ ส่ง Flag กลับ (แก้ไขข้อ 4 ส่วนที่ 3)
        update_flag_in_firebase(ticket_id)
    
    input("\nกด Enter เพื่อปิดโปรแกรม...")
