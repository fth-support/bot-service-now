import time
import pyrebase
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ==========================================
# ส่วนที่ 1: ตั้งค่า Firebase ของจริง
# ==========================================
firebase_config = {
    "apiKey": "AIzaSyCKhox_fzBzuU9n6Q4hv2fHSiXVUw7_I1U",
    "authDomain": "service-now-79151.firebaseapp.com",
    "projectId": "service-now-79151",
    "databaseURL": "https://service-now-79151-default-rtdb.asia-southeast1.firebasedatabase.app/", # ใส่ URL ที่คุณหามา
    "storageBucket": "service-now-79151.firebasestorage.app",
    "messagingSenderId": "437301948553",
    "appId": "1:437301948553:web:3f383db66adf1fababd658"
}

firebase = pyrebase.initialize_app(firebase_config)
db = firebase.database()

def get_new_ticket_from_firebase():
    """ ดึงข้อมูลจริงจาก Firebase เฉพาะตัวที่ status เป็น 'new' """
    print("\n🔍 กำลังตรวจสอบงานใหม่จาก Firebase...")
    try:
        # ดึงข้อมูลจาก node ที่ชื่อ 'tickets' ที่มี status == 'new'
        all_tickets = db.child("tickets").order_by_child("status").equal_to("new").get()
        
        if all_tickets.val():
            # หยิบงานแรกมาทำ
            for ticket in all_tickets.each():
                ticket_id = ticket.key()
                ticket_data = ticket.val()
                return ticket_id, ticket_data
        
        return None, None
    except Exception as e:
        print(f"❌ Error Firebase: {e}")
        return None, None

def update_flag_in_firebase(ticket_id):
    """ เปลี่ยน status เป็น 'completed' เพื่อไม่ให้บอทหยิบมาทำซ้ำ """
    try:
        db.child("tickets").child(ticket_id).update({"status": "completed"})
        print(f"✅ [Firebase] อัปเดตงาน {ticket_id} เป็น 'completed' แล้ว")
    except Exception as e:
        print(f"❌ ไม่สามารถอัปเดตสถานะใน Firebase: {e}")

# ==========================================
# ส่วนที่ 2: ควบคุม ServiceNow (เหมือนเดิม)
# ==========================================
def fill_servicenow_ticket(driver, wait, data):
    try:
        print(f"🚀 เริ่มสร้าง Ticket: {data.get('short_description', 'No Subject')}")
        driver.get("https://keris.service-now.com/incident.do?sys_id=-1")
        time.sleep(3) 

        # เลือก Sense & Respond Demand
        try:
            demand_select = Select(wait.until(EC.presence_of_element_located((By.ID, "incident.u_sense_respond_demand"))))
            demand_select.select_by_visible_text("Retail Solution & Delivery")
        except: pass

        # กรอก Caller
        caller_field = wait.until(EC.element_to_be_clickable((By.ID, "sys_display.incident.caller_id")))
        caller_field.clear()
        caller_field.send_keys(data["caller"])
        time.sleep(1.5)
        caller_field.send_keys(Keys.RETURN)
        
        # กรอกรายละเอียดที่ดึงมาจาก Firebase
        driver.find_element(By.ID, "incident.short_description").send_keys(data["short_description"])
        driver.find_element(By.ID, "incident.description").send_keys(data["description"])
        
        # เลือกค่าอื่นๆ ตาม Fix value
        Select(driver.find_element(By.ID, "incident.category")).select_by_visible_text(data["category"])
        Select(driver.find_element(By.ID, "incident.impact")).select_by_visible_text(data["impact"])
        Select(driver.find_element(By.ID, "incident.urgency")).select_by_visible_text(data["urgency"])
        Select(driver.find_element(By.ID, "incident.contact_type")).select_by_visible_text(data["channel"])
        Select(driver.find_element(By.ID, "incident.state")).select_by_visible_text(data["state"])
        
        # Assignment Group
        ag_field = driver.find_element(By.ID, "sys_display.incident.assignment_group")
        ag_field.clear()
        ag_field.send_keys(data["assignment_group"])
        time.sleep(1.5)
        ag_field.send_keys(Keys.RETURN)

        # Mitigate SLA Description (ใช้ contact_info จาก Firebase)
        driver.find_element(By.ID, "incident.u_mitigate_sla_description").send_keys(data["contact_info"])

        # กด Save
        save_button = driver.find_element(By.ID, "sysverb_insert")
        save_button.click()
        
        time.sleep(4) 
        return True 

    except Exception as e:
        print(f"❌ Error on ServiceNow: {e}")
        return False

# ==========================================
# Main Loop
# ==========================================
if __name__ == "__main__":
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    
    try:
        driver = webdriver.Chrome(options=chrome_options)
        wait = WebDriverWait(driver, 15)
        print("🤖 บอทพร้อมทำงานและกำลังเฝ้า Firebase...")

        while True:
            t_id, t_data = get_new_ticket_from_firebase()
            
            if t_id:
                if fill_servicenow_ticket(driver, wait, t_data):
                    update_flag_in_firebase(t_id)
            else:
                print(".", end="", flush=True)
            
            time.sleep(10) # รอ 10 วินาทีเพื่อเช็คใหม่

    except Exception as e:
        print(f"❌ ไม่สามารถเริ่มบอทได้: {e}")
        input("กด Enter เพื่อปิด...")
