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
    """ 
    จำลองการเช็คงานใหม่ 
    ในอนาคต: ถ้าไม่มีงานใหม่ ให้ฟังก์ชันนี้ return None, None
    """
    # สมมติเงื่อนไข: ถ้าไม่มีข้อมูลใหม่ให้ return (None, None)
    # แต่ตอนนี้ขอจำลองว่ามีงานมา 1 งานเพื่อการทดสอบ
    has_new_data = True 
    
    if not has_new_data:
        return None, None

    ticket_id = "FB_REC_" + str(int(time.time())) # สร้าง ID จำลองตามเวลา
    mock_data = {
        "caller": "MALL-0006", 
        "category": "Software",
        "impact": "5 - Minor",
        "urgency": "5 - Minor", 
        "short_description": "ทดสอบ Loop: ปัญหาระบบ POS", # แก้ไขตามข้อ 1
        "description": "รายละเอียดจาก Firebase...", # แก้ไขตามข้อ 2
        "contact_info": "คุณใหญ่ 0926297894", # แก้ไขตามข้อ 3
        "channel": "Phone",
        "state": "Acknowledged",
        "assignment_group": "FTH ISS3000 MALL",
        "reported_by": "DON-001"
    }
    return ticket_id, mock_data

def update_flag_in_firebase(ticket_id):
    """ ส่ง Flag กลับไปบอก Firebase (ข้อ 4 ส่วนที่ 3) """
    print(f"✅ [Firebase] อัปเดตสถานะงาน {ticket_id} เป็น 'Completed' เรียบร้อย!")
    # โค้ดจริง: db.child("tickets").child(ticket_id).update({"status": "completed"})

# ==========================================
# ส่วนที่ 2: ควบคุมหน้าเว็บ ServiceNow
# ==========================================
def fill_servicenow_ticket(driver, wait, data):
    try:
        print("\n🚀 กำลังเริ่มสร้าง Ticket ใหม่...")
        raw_form_url = "https://keris.service-now.com/incident.do?sys_id=-1"
        driver.get(raw_form_url)
        time.sleep(3) 

        # --- ส่วนที่ 2 ข้อ 4: เลือก Sense & Respond Demand ---
        try:
            demand_select = Select(wait.until(EC.presence_of_element_located((By.ID, "incident.u_sense_respond_demand"))))
            demand_select.select_by_visible_text("Retail Solution & Delivery")
        except:
            print("⚠️ ไม่พบฟิลด์ Sense & Respond Demand (ข้ามสเต็ปนี้)")

        # กรอก Caller
        caller_field = wait.until(EC.element_to_be_clickable((By.ID, "sys_display.incident.caller_id")))
        caller_field.clear()
        caller_field.send_keys(data["caller"])
        time.sleep(1.5)
        caller_field.send_keys(Keys.RETURN)
        
        # เลือก Category, Impact, Urgency
        Select(driver.find_element(By.ID, "incident.category")).select_by_visible_text(data["category"])
        Select(driver.find_element(By.ID, "incident.impact")).select_by_visible_text(data["impact"])
        Select(driver.find_element(By.ID, "incident.urgency")).select_by_visible_text(data["urgency"])

        # --- ข้อ 1 & 2: Mapping ฟิลด์จาก Firebase ---
        driver.find_element(By.ID, "incident.short_description").send_keys(data["short_description"])
        driver.find_element(By.ID, "incident.description").send_keys(data["description"])

        Select(driver.find_element(By.ID, "incident.contact_type")).select_by_visible_text(data["channel"])
        Select(driver.find_element(By.ID, "incident.state")).select_by_visible_text(data["state"])
        
        # Assignment Group
        ag_field = driver.find_element(By.ID, "sys_display.incident.assignment_group")
        ag_field.clear()
        ag_field.send_keys(data["assignment_group"])
        time.sleep(1.5)
        ag_field.send_keys(Keys.RETURN)

        # Reported By
        rep_field = driver.find_element(By.ID, "sys_display.incident.u_reported_by")
        rep_field.clear()
        rep_field.send_keys(data["reported_by"])
        time.sleep(1.5)
        rep_field.send_keys(Keys.RETURN)

        # --- ข้อ 3: Mitigate SLA (ใช้ contact_info) ---
        driver.find_element(By.ID, "incident.u_mitigate_sla_description").send_keys(data["contact_info"])

        # --- ส่วนที่ 3 ข้อ 5: กด Save ---
        print("💾 กำลังกด Save...")
        save_button = driver.find_element(By.ID, "sysverb_insert")
        save_button.click()
        
        time.sleep(3) # รอให้ระบบบันทึกและเปลี่ยนหน้า
        print("🎉 บันทึก Ticket สำเร็จ!")
        return True 

    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดระหว่างกรอกข้อมูล: {e}")
        return False

# ==========================================
# Main: ระบบ Loop ตรวจสอบงาน
# ==========================================
if __name__ == "__main__":
    print("🤖 Bot is starting... (Press Ctrl+C to stop)")
    
    # ตั้งค่า Chrome ครั้งเดียวที่นอก Loop
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    
    try:
        driver = webdriver.Chrome(options=chrome_options)
        wait = WebDriverWait(driver, 15)
        
        # --- เริ่มต้น Loop การทำงาน ---
        while True:
            # 1. เช็ค Firebase ว่ามีงานใหม่ไหม
            ticket_id, ticket_data = get_new_ticket_from_firebase()
            
            if ticket_id and ticket_data:
                # 2. ถ้ามีงาน ให้เริ่มกรอก
                if fill_servicenow_ticket(driver, wait, ticket_data):
                    # 3. ถ้ากรอกและ Save สำเร็จ ให้ส่ง Flag กลับ Firebase
                    update_flag_in_firebase(ticket_id)
                
                print("\n✅ งานเสร็จเรียบร้อย กำลังรอตรวจสอบงานใหม่ต่อไป...")
            else:
                # ถ้าไม่มีงานใหม่ ให้พิมพ์จุดบอกสถานะ (เพื่อไม่ให้หน้าจอนิ่งเกินไป)
                print(".", end="", flush=True)
            
            # 4. พัก 10 วินาที ก่อนจะวนไปเช็ค Firebase อีกครั้ง (เพื่อไม่ให้ CPU ทำงานหนักเกินไป)
            time.sleep(10)

    except Exception as e:
        print(f"\n❌ ไม่สามารถเชื่อมต่อกับ Chrome ได้: {e}")
        print("กรุณาเปิด Chrome ด้วยโหมด Debugging (Port 9222) ก่อนรันบอท")
        input("กด Enter เพื่อปิด...")
