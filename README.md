# 🚛 Trailer Status Dashboard

ระบบตรวจสอบสถานะตู้รถอัตโนมัติ สำหรับ Lotus / Linfox Fleet Management

## วิธีติดตั้งและรัน

### 1. ติดตั้ง dependencies
```bash
pip install -r requirements.txt
```

### 2. รัน Streamlit app
```bash
streamlit run app.py
```

จากนั้นเปิด browser ที่ `http://localhost:8501`

---

## การใช้งาน

### ไฟล์ที่ต้องอัพโหลด (Excel .xlsx / .xlsm)

| ไฟล์ | รายละเอียด | Column ที่ใช้ |
|---|---|---|
| **Trailer Lotus** | รายการป้ายทะเบียนตู้รถทั้งหมด | `Fleet#` |
| **VOR Allnow** | ตู้ชำรุด + รอขาย | `Fleet No` (ทุก sheet) |
| **VOR Linfox** | ตู้ชำรุด + รอขาย | `Fleet No` (sheet วันที่ + Disposal Status) |
| **Trucker** | ข้อมูล Trip | `Trailer`, `Confirmed Depart DC Time`, `Confirmed Return DC Time`, `Depart Depot` |

---

## ลำดับการตรวจสอบสถานะ

```
Trailer Lotus (ทะเบียนทั้งหมด)
        │
        ▼
 [1] VOR Allnow
     - sheet VOR report  → ถ้าเจอ = VOR
     - sheet รอขาย       → ถ้าเจอ = VOR
        │ ไม่เจอ
        ▼
 [2] VOR Linfox
     - sheet วันที่ (เช่น "10 Jun 26") → ถ้าเจอ = VOR
     - sheet Disposal Status           → ถ้าเจอ = VOR
        │ ไม่เจอ
        ▼
 [3] Trucker (ดูจากล่างขึ้นบน)
     3.1 วันเดียวกับวันนี้ + ไม่มี Return → ชื่อ Depart Depot
     3.2 ไม่ใช่วันนี้   + มี Return       → Parking + ชื่อ Depot
     3.3 ไม่ใช่วันนี้   + ไม่มี Return    → On Road
        │ ไม่เจอ
        ▼
     N/A
```

---

## สถานะที่แสดง

| สถานะ | สี | ความหมาย |
|---|---|---|
| ⚠️ VOR | แดง | ตู้ชำรุด / รอขาย |
| 🏭 ชื่อ Depot | น้ำเงิน | ออกเดินทางวันนี้ ยังไม่กลับ |
| 🅿️ Parking + Depot | เหลือง | กลับ Depot แล้ว |
| 🚛 On Road | เขียว | อยู่ระหว่างเส้นทาง |
| ⚪ N/A | เทา | ไม่พบข้อมูล |

---

## Feature

- ✅ ตรวจสอบสถานะตู้รถอัตโนมัติจาก 4 ไฟล์ Excel
- ✅ แสดงตาราง Trailer Lotus ทั้งหมด พร้อม column Status
- ✅ Summary metrics (นับจำนวนแต่ละสถานะ)
- ✅ ค้นหาป้ายทะเบียน + กรองสถานะ
- ✅ Download ผลลัพธ์เป็น Excel
- ✅ เลือกวันที่ประมวลผลได้

---

*Trailer Status Dashboard v1.0*
