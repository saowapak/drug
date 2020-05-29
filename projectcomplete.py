from openpyxl import Workbook #import package ที่ชื่อว่า openpyxl ก็คือให้โปรแกรมทำงานร่วมกับโปรแกรม Excel
import _datetime # import เวลา เข้ามา


# สร้าง dict เพื่อเก็บข้อมูลยา
drugdata = {"ยาบรรเทาอาการปวด/ลดไข้": ["ยาน้ำTempra kid" ,"Tempra Forte" , "ยาพาราเซตามอล"], #โดยแบ่งประเภทของยาเป็นดังนี้
            "ยาแก้ปวดท้อง": ["Berclomine","ยาธาตุน้ำขาว ALUM MILK","Buscopan"],
            "ยาแก้อาเจียน": ["Molax siam "],
            "น้ำเกลือแร่": ["ZEDAN siam"],
            "ยาทาแก้แมลงสัตว์กัดต่อย":["ZAM BUK ฝาเขียว"],
            "ยาแก้ปวดประจำเดือน": ["Oreda ORS Powder"]
            }

excelfile = Workbook()  # สร้างไฟล์ excel ใน python
data = excelfile.active  # เลือก worksheet ที่กำลังเปิดอยู่ โดยให้เก็บไว้ในตัวแปร data
data.append(['บันทึกการจ่ายยาของห้องพยาบาล'])  # printคำว่า บันทึกการจ่ายยาของห้องพยาบาล ในตัวแปรที่ชื่อว่า data เพื่อเป็นหัวตาราง
data.append(['ชื่อ-สกุล','ชั้น', 'ยา', 'อาการ', 'จำนวน', 'เวลา'])# printคำว่า ชื่อ-สกุล,ยา,อาการ,จำนวน,เวลา ในตัวแปรที่ชื่อว่า data #โดยข้อมูลที่ปริ้นออกมาจะแบ่งเป็นช่วง ดังนี้

a = []

while True: 
    manu = input('ขอยา[a],ออกจากระบบ[q]\n')  #สร้างตัวแปร manu เพื่อรับค่าว่าผู้ใช้จะขอยาหรือออกจากระบบ
    manu = manu.lower () 
    if manu == 'a': # manu มีค่า = a ให้ทำตามคำสั่งดังนี้
        name = input('ชื่อ-สกุล : ') #สร้างตัวแปร name เพื่อเก็บชื่อของผู้ใช้
        grade = input('ชั้น : ')
        print('ยาบรรเทาอาการปวด/ลดไข้\n 1.ยาน้ำTempra kid\n 2.Tempra Forte\n 3.ยาพาราเซตามอล\n')
        print('ยาแก้ปวดท้อง\n 4.Berclomine\n 5.ยาธาตุน้ำขาว ALUM MILK\n 6.Buscopan\n')
        print('ยาแก้อาเจียน\n 7.Molax siam\n')
        print('น้ำเกลือแร่ \n 8.ZEDAN  siam\n')
        print('ยาทาแก้แมลงสัตว์กัดต่อย\n 9.ZAM BUK ฝาเขียว\n')
        print('ยาแก้ปวดประจำเดือน\n 10.Oreda ORS Powder\n')#แสดงรายการยา พร้อมหมายเลขยาให้ผู้ใช้เลือก
        drug = input('ยา : ')  # สร้างตัวแปร drug เพื่อเก็บหมายเลขยาที่ผู้ใช้เลือก
        drugname = []  # สร้างตัวแปร drugname เพื่อเก็บชื่อยา
        if drug == '1': #ถ้าผู้ใช้เลือกยาหมายเลข 1
            drugname = (drugdata["ยาบรรเทาอาการปวด/ลดไข้"][0]) # สร้างตัวแปร drugname เพื่อเก็บชื่อยาโดยจะเลือกชื่อยาที่ 0 ในหมวด ยาบรรเทาอาการปวด/ลดไข้ ใน dict drugdata
        elif drug =='2':
            drugname = (drugdata["ยาบรรเทาอาการปวด/ลดไข้"][1])
        elif drug == '3':
            drugname = (drugdata["ยาบรรเทาอาการปวด/ลดไข้"][2])
        elif drug == '4':
            drugname = (drugdata["ยาแก้ปวดท้อง"][0])
        elif drug == '5':
            drugname = (drugdata["ยาแก้ปวดท้อง"][1])
        elif drug == '6':
            drugname = (drugdata["ยาแก้ปวดท้อง"][2])
        elif drug == '7':
            drugname = (drugdata["ยาแก้อาเจียน"][0])
        elif drug == '8':
            drugname = (drugdata["น้ำเกลือแร่"][0])
        elif drug == '9':
            drugname = (drugdata["ยาทาแก้แมลงสัตว์กัดต่อย"][0])
        elif drug == '10':
            drugname = (drugdata["ยาแก้ปวดประจำเดือน"][0])
        elif drug != range(1,10): #ถ้าผู้ใช้กรอกหมายเลข ที่ไม่อยู่ในช่วง 1-10 ก็จะให้โปรแกรมเริ่มใหม่
            continue
        symptom = input('อาการ : ') # สร้างตัวแปร symtom เพื่อเก็บข้อมูลอาการ
        amout = input('จำนวนยาที่ต้องการ(เม็ด/ช้อน): ')# สร้างตัวแปร amout เพื่อเก็บข้อมูลจำนวนยาที่ต้องการ
        print('ข้อมูลได้บันทึกเรียบร้อย')
        time = _datetime.datetime.now()  # สร้างตัวแปร time เพื่อเก็บข้อมูลเวลาปัจจุบัน
        data.append([name,grade,drugname, symptom, amout, time])#เพิ่มข้อมูล name,drugname, symptom, amout, time ใน data
        excelfile.save('result.xlsx')#save file excel โดยให้ชื่อว่า result

    elif manu == 'q':  # manu มีค่า = q ให้ออกจาdโปรแกรม
        break



