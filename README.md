 ' =========================================================
    ' ---- 8. СОХРАНЕНИЕ 4 ВКЛАДОК В ОТДЕЛЬНЫЕ ФАЙЛЫ ----
    ' =========================================================
    Dim desktopPath As String
    Dim folderPath As String
    Dim savePath As String
    Dim wbNew As Workbook
    
    ' Путь к рабочему столу
    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    
    ' Создаем папку с датой и временем
    folderPath = desktopPath & "\" & ChrW(1042) & ChrW(1099) & ChrW(1075) & ChrW(1088) & ChrW(1091) & ChrW(1079) & ChrW(1082) & ChrW(1072) & "_" & Format(Now, "ddmmyyyy_hhmmss")
    
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 1. Сохраняем СДУ_ТО_цены_для_загрузки
    wsSDU.Copy
    Set wbNew = ActiveWorkbook
    savePath = folderPath & "\" & ChrW(1057) & ChrW(1044) & ChrW(1059) & "_" & ChrW(1058) & ChrW(1054) & "_" & ChrW(1094) & ChrW(1077) & ChrW(1085) & ChrW(1099) & "_" & ChrW(1076) & ChrW(1083) & ChrW(1103) & "_" & ChrW(1079) & ChrW(1072) & ChrW(1075) & ChrW(1088) & ChrW(1091) & ChrW(1079) & ChrW(1082) & ChrW(1080) & ".xlsx"
    wbNew.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    wbNew.Close False
    
    ' 2. Сохраняем СДУ_переодиность_ТО
    wsSDUPeriod.Copy
    Set wbNew = ActiveWorkbook
    savePath = folderPath & "\" & ChrW(1057) & ChrW(1044) & ChrW(1059) & "_" & ChrW(1087) & ChrW(1077) & ChrW(1088) & ChrW(1077) & ChrW(1086) & ChrW(1076) & ChrW(1080) & ChrW(1085) & ChrW(1086) & ChrW(1089) & ChrW(1090) & ChrW(1100) & "_" & ChrW(1058) & ChrW(1054) & ".xlsx"
    wbNew.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    wbNew.Close False
    
    ' 3. Сохраняем NAV_справочник_цен
    wsNAV.Copy
    Set wbNew = ActiveWorkbook
    savePath = folderPath & "\NAV_" & ChrW(1089) & ChrW(1087) & ChrW(1088) & ChrW(1072) & ChrW(1074) & ChrW(1086) & ChrW(1095) & ChrW(1085) & ChrW(1080) & ChrW(1082) & "_" & ChrW(1094) & ChrW(1077) & ChrW(1085) & ".xlsx"
    wbNew.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    wbNew.Close False
    
    ' 4. Сохраняем NAV_Справочник_переодичности_ТО
    wsNAVPeriod.Copy
    Set wbNew = ActiveWorkbook
    savePath = folderPath & "\NAV_" & ChrW(1057) & ChrW(1087) & ChrW(1088) & ChrW(1072) & ChrW(1074) & ChrW(1086) & ChrW(1095) & ChrW(1085) & ChrW(1080) & ChrW(1082) & "_" & ChrW(1087) & ChrW(1077) & ChrW(1088) & ChrW(1077) & ChrW(1086) & ChrW(1076) & ChrW(1080) & ChrW(1095) & ChrW(1085) & ChrW(1086) & ChrW(1089) & ChrW(1090) & ChrW(1080) & "_" & ChrW(1058) & ChrW(1054) & ".xlsx"
    wbNew.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    wbNew.Close False
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' ---- Итоговое сообщение ----
    MsgBox ChrW(1043) & ChrW(1086) & ChrW(1090) & ChrW(1086) & ChrW(1074) & ChrW(1086) & "!" & vbCrLf & _
           ChrW(1060) & ChrW(1072) & ChrW(1081) & ChrW(1083) & ChrW(1099) & " " & ChrW(1089) & ChrW(1086) & ChrW(1093) & ChrW(1088) & ChrW(1072) & ChrW(1085) & ChrW(1077) & ChrW(1085) & ChrW(1099) & " " & ChrW(1074) & " " & folderPath, vbInformation
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Ошибка: " & Err.Description, vbCritical
End Sub
