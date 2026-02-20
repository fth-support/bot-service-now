from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

def fill_servicenow_ticket():
    print("กำลังเชื่อมต่อกับ Chrome ที่เปิดอยู่ (Port 9222)...")
    
    # 1. ตั้งค่าให้ Selenium เชื่อมต่อไปยัง Chrome ที่เปิดรอไว้
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    
    try:
        driver = webdriver.Chrome(options=chrome_options)
        print("เชื่อมต่อสำเร็จ! กำลังเริ่มทำงาน...")
        wait = WebDriverWait(driver, 15)

        # ข้อมูลจำลอง (Mock Data) อิงจากที่คุณระบุมา 12 ข้อ
        # อนาคตเราจะเปลี่ยนตรงนี้ให้ไปดึงจาก Firebase แทนครับ
        data = {
            "caller": "MALL-0006", 
            "category": "Software",
            "impact": "5 - Minor",  # ค่า Dropdown อาจจะเป็นแค่เลข 5 หรือข้อความเต็ม ต้องลองเทสดูครับ
            "urgency": "5 - Minor", 
            "short_desc": "ทดสอบระบบ Auto Fill จาก Bot",
            "desc": "นี่คือการทดสอบกรอกข้อมูลอัตโนมัติ\nบรรทัดที่สอง",
            "channel": "Phone",
            "state": "Acknowledged",
            "assignment_group": "FTH ISS3000 MALL",
            "reported_by": "DON-001",
            "mitigate_sla": "คุณใหญ่ 0926297894"
        }

        # 2. จำลองการกดปุ่ม New (ID ของปุ่มใน List view มักจะเป็น sysverb_new)
        print("กำลังกดปุ่ม New...")
        new_button = wait.until(EC.element_to_be_clickable((By.ID, "sysverb_new")))
        new_button.click()

        # 3. เข้าสู่ iframe (สำคัญมาก! ฟอร์ม ServiceNow มักจะอยู่ในกรอบที่ชื่อ gsft_main)
        print("รอโหลดหน้าฟอร์ม...")
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "gsft_main")))

        # 4. เริ่มกรอกข้อมูล
        print("กำลังกรอกข้อมูล...")

        # ข้อ 1: Caller (พิมพ์แล้วรอ 1.5 วิ ให้ระบบค้นหา แล้วกด Enter)
        caller_field = wait.until(EC.element_to_be_clickable((By.ID, "sys_display.incident.caller_id")))
        caller_field.clear()
        caller_field.send_keys(data["caller"])
        time.sleep(1.5) 
        caller_field.send_keys(Keys.RETURN)

        # ข้อ 2: Category (Dropdown)
        Select(driver.find_element(By.ID, "incident.category")).select_by_visible_text(data["category"])

        # ข้อ 3, 4: Impact & Urgency (Dropdown)
        Select(driver.find_element(By.ID, "incident.impact")).select_by_visible_text(data["impact"])
        Select(driver.find_element(By.ID, "incident.urgency")).select_by_visible_text(data["urgency"])

        # ข้อ 6, 7: Short description & Description
        driver.find_element(By.ID, "incident.short_description").send_keys(data["short_desc"])
        driver.find_element(By.ID, "incident.description").send_keys(data["desc"])

        # ข้อ 8: Channel (ใน SNOW มักใช้ชื่อ contact_type)
        Select(driver.find_element(By.ID, "incident.contact_type")).select_by_visible_text(data["channel"])

        # ข้อ 9: State 
        Select(driver.find_element(By.ID, "incident.state")).select_by_visible_text(data["state"])

        # ข้อ 10: Assignment group (พิมพ์แล้วกด Enter)
        ag_field = driver.find_element(By.ID, "sys_display.incident.assignment_group")
        ag_field.clear()
        ag_field.send_keys(data["assignment_group"])
        time.sleep(1.5)
        ag_field.send_keys(Keys.RETURN)

        # ⚠️ ข้อ 11, 12: เป็นฟิลด์ Custom (ID มักจะขึ้นต้นด้วย u_)
        # บรรทัดนี้อาจจะ Error ได้ถ้า ID ไม่ตรง คุณต้องเปิด F12 ไปดู ID ของช่องนี้อีกทีครับ
        # ขอใส่ ID สมมติไว้ก่อนนะครับ (u_reported_by และ u_mitigate_sla_description)
        
        # rep_field = driver.find_element(By.ID, "sys_display.incident.u_reported_by")
        # rep_field.clear()
        # rep_field.send_keys(data["reported_by"])
        # time.sleep(1.5)
        # rep_field.send_keys(Keys.RETURN)

        # driver.find_element(By.ID, "incident.u_mitigate_sla_description").send_keys(data["mitigate_sla"])

        print("กรอกข้อมูลเสร็จสิ้น! (ยังไม่ได้สั่งกด Save เพื่อให้คุณตรวจสอบความถูกต้องหน้าจอก่อน)")
        
        # ออกจาก iframe กลับสู่หน้าจอหลัก
        driver.switch_to.default_content()

    except Exception as e:
        print(f"เกิดข้อผิดพลาดในการทำงาน: {e}")

if __name__ == "__main__":
    fill_servicenow_ticket()
    input("กด Enter เพื่อปิดหน้าต่างนี้...")
