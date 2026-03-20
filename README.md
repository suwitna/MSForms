# 🚀 MS Forms to SQL Express via OneDrive Sync

ตัวอย่าง Data Pipeline แบบ Low-code โดยใช้ Microsoft Forms เป็นตัวรับข้อมูล, OneDrive เป็นตัวกลางในการ Sync ไฟล์ Excel ลงเครื่อง PC และใช้ PowerShell ในการ ETL ข้อมูลเข้าสู่ MS SQL Server Express


<img width="1784" height="767" alt="image" src="https://github.com/user-attachments/assets/7172bf09-63c7-4372-87be-d01f99ef69f0" />

-----

## 📂 1. การสร้างไฟล์ Excel เชื่อมกับ MS Forms

วิธีการสร้างจะแตกต่างกันตามประเภทบัญชีที่คุณใช้งาน ดังนี้:

### 🏢 สำหรับบัญชีองค์กร (Microsoft 365 Business/Education)

วิธีนี้เสถียรที่สุด ข้อมูลจะถูกเขียนลง Excel Real-time

1.  Log-in เข้าสู่ [OneDrive for Business](https://www.google.com/search?q=https://portal.office.com).
2.  คลิก **+ New (สร้างใหม่)** \> เลือก **Forms for Excel**.
3.  ตั้งชื่อไฟล์ Excel (เช่น `Feedback_Data.xlsx`).
4.  เบราว์เซอร์จะเปิดหน้า Microsoft Forms ให้คุณออกแบบคำถาม.
5.  **ผลลัพธ์:** ไฟล์ Excel จะถูกสร้างใน OneDrive และมีสัญลักษณ์รูป Forms อยู่ที่ไอคอนไฟล์.

### 🏠 สำหรับบัญชีบุคคล (Personal - @outlook.com / @hotmail.com)

ต้องสร้างไฟล์ Excel ก่อนแล้วค่อย "ฝัง" Form ลงไป

1.  Log-in เข้าสู่ [OneDrive.com](https://onedrive.live.com).

2.  คลิก **+ New (สร้างใหม่)** \> **Excel workbook**. 
<img width="661" height="563" alt="image" src="https://github.com/user-attachments/assets/93b3d26f-055a-4c52-ab78-75d7d177f633" />

3.  เมื่อไฟล์เปิดขึ้นมา ให้ไปที่แถบเมนู **Insert (แทรก)** \> คลิกปุ่ม **Forms** \> **+ New Form**.
<img width="716" height="405" alt="image" src="https://github.com/user-attachments/assets/5d0f19f8-ee52-48e7-b3d9-a86deaaf98d4" />

4.  ออกแบบคำถามในหน้าต่างที่ปรากฏขึ้น.
<img width="980" height="800" alt="image" src="https://github.com/user-attachments/assets/7ef3e80b-25ee-403b-9b62-d220de69061f" />

5.  **ผลลัพธ์:** ข้อมูลจาก Form จะถูกส่งมาที่ Sheet ใหม่ (ปกติชื่อ *Form1*) ในไฟล์ Excel นี้.
<img width="1111" height="160" alt="image" src="https://github.com/user-attachments/assets/6f9c3227-409e-4844-b8e1-1a7fc9a481ce" />

-----

## 🔄 2. การ Sync ไฟล์ลงเครื่อง PC

เพื่อให้ PowerShell อ่านไฟล์ได้ คุณต้อง Sync ไฟล์จาก Cloud ลงมาที่ Disk:

1.  เปิดโปรแกรม **OneDrive** บนคอมพิวเตอร์และ Log-in บัญชีเดียวกับที่สร้างไฟล์.
<img width="1032" height="359" alt="image" src="https://github.com/user-attachments/assets/aa2552e0-e7df-46fa-b07a-9ebbb3a99689" />

2.  รอให้ไฟล์ Excel ปรากฏใน Folder OneDrive บนเครื่อง.
<img width="847" height="349" alt="image" src="https://github.com/user-attachments/assets/5d96ba83-0e79-4b35-b45f-2f821dce32c2" />

3.  **สำคัญ:** คลิกขวาที่ไฟล์นั้น แล้วเลือก **"Always keep on this device"** เพื่อให้ไฟล์พร้อมใช้งานแบบ Offline เสมอ และ PowerShell สามารถเข้าถึงได้ตลอดเวลา.
<img width="746" height="335" alt="image" src="https://github.com/user-attachments/assets/1d9a3215-aa52-4950-8eaf-60d66a5b0882" />

-----

## 🛠️ 3. PowerShell Script: MS Forms to SQL Express

*   คำสั่ง Import-Excel ซึ่งไม่ได้ติดมากับ Windows ต้องลงเพิ่มผ่าน PowerShell (Run as Administrator):

```powershell
# ติดตั้ง Module สำหรับอ่านไฟล์ Excel โดยไม่ต้องมีโปรแกรม Excel ในเครื่อง
Install-Module -Name ImportExcel -Force -Scope CurrentUser
```

```powershell
# --- ตั้งค่าเครื่องและไฟล์ ---
$excelPath = "D:\Suwit\OneDrive\SurveyDataForms.xlsx" 
$tempPath = "$env:TEMP\temp_forms_data.xlsx" # ไฟล์สำรองชั่วคราวในโฟลเดอร์ Temp ของ Windows
$serverName = ".\SQLEXPRESS"
$dbName = "SADB"
$connString = "Server=$serverName;Database=$dbName;Integrated Security=True;Encrypt=True;TrustServerCertificate=True;"

try {
    # --- ขั้นตอนกันเหนียว: Copy ไฟล์ออกมาแม้ไฟล์จะถูกเปิดค้างไว้ ---
    # ใช้ -Force เพื่อทับไฟล์เก่า และเผื่อไฟล์ต้นทางถูก Lock (Read-only copy)
    Copy-Item -Path $excelPath -Destination $tempPath -Force -ErrorAction Stop
    Write-Host "Copying Excel to temp location..." -ForegroundColor Cyan

    # --- 1. อ่านข้อมูลจากไฟล์ Temp ---
    $data = Import-Excel -Path $tempPath

    # --- 2. เตรียมเชื่อมต่อ SQL ---
    $connection = New-Object System.Data.SqlClient.SqlConnection($connString)
    $connection.Open()

    foreach ($row in $data) {
        # กันเหนียว: เช็คว่าแถวนั้นมี Id จริงไหม (ป้องกันแถวว่างท้ายไฟล์)
        if ([string]::IsNullOrWhiteSpace($row.Id)) { continue }

        # --- 3. ตรวจสอบข้อมูลซ้ำ ---
        $checkCmd = $connection.CreateCommand()
        $checkCmd.CommandText = "SELECT COUNT(*) FROM SurveyResponses WHERE Id = @id"
        $checkCmd.Parameters.AddWithValue("@id", $row.Id)
        $exists = $checkCmd.ExecuteScalar()

        if ($exists -eq 0) {
            # --- 4. ถ้ายังไม่มี ให้ Insert ---
            $insertQuery = @"
            INSERT INTO SurveyResponses (Id, StartTime, CompletionTime, Email, FullName, Question1, Question2, Question3)
            VALUES (@id, @start, @complete, @email, @name, @q1, @q2, @q3)
"@
            $cmd = $connection.CreateCommand()
            $cmd.CommandText = $insertQuery
            
            # ใช้ Try-Catch ย่อยเพื่อป้องกันข้อมูลบางแถวผิดพลาดแล้วทำให้ Script หยุดรันทั้งหมด
            try {
                $cmd.Parameters.AddWithValue("@id", $row.Id)
                $cmd.Parameters.AddWithValue("@start", $row."Start time")
                $cmd.Parameters.AddWithValue("@complete", $row."Completion time")
                $cmd.Parameters.AddWithValue("@email", ([string]$row.Email))
                $cmd.Parameters.AddWithValue("@name", ([string]$row.Name))
                $cmd.Parameters.AddWithValue("@q1", ([string]$row.Question1))
                $cmd.Parameters.AddWithValue("@q2", ([string]$row.Question2))
                $cmd.Parameters.AddWithValue("@q3", ([string]$row.Question3))
                
                $cmd.ExecuteNonQuery()
                Write-Host "Inserted Id: $($row.Id) successfully." -ForegroundColor Green
            } catch {
                Write-Host "Failed to insert Id: $($row.Id). Error: $($_.Exception.Message)" -ForegroundColor Red
            }
        } else {
            Write-Host "Id: $($row.Id) already exists. Skipping..." -ForegroundColor Yellow
        }
    }

    $connection.Close()
}
catch {
    Write-Host "Critical Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    # ลบไฟล์ Temp ทิ้งเมื่อเสร็จงาน (เพื่อความสะอาด)
    if (Test-Path $tempPath) { Remove-Item $tempPath -Force }
}
```
---

### 📝 คำอธิบายเพิ่มเติม

1.  **Column Mapping:** ในสคริปต์มีการใช้ `$row.'Start time'` (ต้องใส่ single quote ครอบชื่อที่มีเว้นวรรค) เพื่อให้อ่านชื่อคอลัมน์ที่ MS Forms สร้างให้อัตโนมัติได้ถูกต้อง
2.  **SQL Data Types:** แนะนำให้สร้างตารางใน SQL Express โดยใช้ Data Type ดังนี้:
    *   `Id`: int (Primary Key)
    *   `StartTime` & `CompletionTime`: datetime
    *   `Email` & `FullName`: nvarchar(255)
    *   `Questions`: nvarchar(max) หรือ nvarchar(500)
3.  **Security:** สคริปต์ใช้ `Integrated Security=True` (Windows Authentication) ซึ่งปลอดภัยและสะดวกสำหรับการใช้งานภายในเครื่องหรือในโดเมนบริษัท

## ⚠️ ข้อควรระวัง (Tips)

  * **OneDrive Sync Latency:** บางครั้งข้อมูลใน Cloud อาจใช้เวลา 10-30 วินาทีกว่าจะ Sync ลง PC.
  * **Column Names:** ชื่อหัวตารางใน Excel จะตรงกับ "คำถาม" ใน MS Forms หากคุณแก้ไขคำถามในภายหลัง ต้องมาแก้ชื่อคอลัมน์ในสคริปต์ PowerShell ด้วย.
  * **SQL Permissions:** ตรวจสอบให้แน่ใจว่า User ที่รัน PowerShell มีสิทธิ์ `db_datareader` และ `db_datawriter` ใน SQL Express.

-----
## 🗄️ 4. SQL Server Schema Setup

```sql
-- สร้างตารางที่แมปตามหัวข้อใน Excel ของ MS Forms
CREATE TABLE SurveyResponses (
    Id INT PRIMARY KEY,              -- ตรงกับคอลัมน์ 'ID'
    StartTime DATETIME,                  -- ตรงกับคอลัมน์ 'Start time'
    CompletionTime DATETIME,             -- ตรงกับคอลัมน์ 'Completion time'
    Email NVARCHAR(255),                 -- ตรงกับคอลัมน์ 'Email'
    Name NVARCHAR(255),                  -- ตรงกับคอลัมน์ 'Name'
    Question1 NVARCHAR(MAX),             -- ตรงกับคอลัมน์ 'Question1'
    Question2 NVARCHAR(MAX),             -- ตรงกับคอลัมน์ 'Question2'
    Question3 NVARCHAR(MAX),             -- ตรงกับคอลัมน์ 'Question3'
    InsertedAt DATETIME DEFAULT GETDATE() -- วันเวลาที่ข้อมูลถูกนำเข้า SQL
);
GO
```

---

## 🚀 วิธีการใช้งาน (Usage)

รัน PowerShell ผ่าน Terminal โดยใส่ Option ตามที่คุณต้องการได้ดังนี้:

```powershell
./Sync-SurveyData.ps1
```



<img width="1074" height="205" alt="image" src="https://github.com/user-attachments/assets/b8b125ee-e175-4fe7-bf3d-9d38ce1d6b46" />



<img width="862" height="465" alt="image" src="https://github.com/user-attachments/assets/1769fe53-5e45-41fa-ac54-5ef8fb23c532" />



<img width="820" height="395" alt="image" src="https://github.com/user-attachments/assets/aba3a107-39e1-4c06-9bdb-b6af97d027d3" />

---

## 💡 สรุป

*   **Zero Locking:** ใช้วิธี Copy ไฟล์ไปที่ Temp ก่อนอ่าน ทำให้ OneDrive สามารถ Sync ได้ตลอดเวลาโดยไม่เกิด Error "File in use"
*   **Data Integrity:** มีระบบเช็ค ID ซ้ำ (NoDup) ทำให้ข้อมูลใน SQL ไม่บวมจากการรันสคริปต์ซ้ำๆ
*   **Flexibility:** รองรับทั้ง Microsoft 365 องค์กร และ OneDrive บุคคล
*   **Scalability:** สามารถนำไปตั้งค่าใน **Windows Task Scheduler** ให้ทำงานอัตโนมัติได้ทุกๆ 15 นาที

---
