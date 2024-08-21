import pyautogui
import openpyxl
import time

# โหลดไฟล์ Excel และเลือกชีท
wb = openpyxl.load_workbook('Book1.xlsx')
sheet = wb['Sheet1']  # เปลี่ยนชื่อชีทตามที่คุณใช้

# รอเวลาสำหรับคุณไปคลิกที่หน้าต่างโปรแกรม EXE
time.sleep(0.5)

# เลือกข้อมูลจากแถวที่ 2 คอลัมน์ Material description
for row in sheet.iter_rows(min_row=2, max_row=2, min_col=2, max_col=2, values_only=True):
    for cell in row:
        if cell:
            pyautogui.click(x=1048, y=230)  
            pyautogui.write(str(cell))   
            pyautogui.click(x=1222, y=239)
