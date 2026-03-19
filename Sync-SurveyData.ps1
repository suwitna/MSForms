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