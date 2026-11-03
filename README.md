    ' ---- 2. Создаём ДИНАМИЧЕСКИЙ лист результата (с датой и временем) ----
    Dim destSheetName As String
    destSheetName = "Результат_" & Format(Now, "ddmmyy_hhmmss")
    
    Set wsDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsDest.Name = destSheetName
