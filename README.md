        ' Сначала проверяем специальные Service_ID для LADA ISKRA
        If ladaIskraServiceIDs.exists(serviceId) Then
            periodStart = 0
            periodEnd = 13
        ' Затем проверяем специальные Service_ID для LADA Vesta
        ElseIf ladaVestaServiceIDs.exists(serviceId) Then
            periodStart = 0
            periodEnd = 13
        ' ГЛАВНОЕ: если марка = SOLLERS
        ElseIf brand = "SOLLERS" Then
            If InStr(1, serviceId, "ST8", vbTextCompare) > 0 Then
                periodStart = 1
                periodEnd = 20
            Else
                Select Case True
                    Case InStr(1, serviceId, "Atlant", vbTextCompare) > 0
                        periodStart = 0
                        periodEnd = 10
                    Case InStr(1, serviceId, "SF4", vbTextCompare) > 0
                        periodStart = 0
                        periodEnd = 10
                    Case InStr(1, serviceId, "SF5", vbTextCompare) > 0
                        periodStart = 0
                        periodEnd = 10
                    Case InStr(1, serviceId, "SF1", vbTextCompare) > 0
                        periodStart = 0
                        periodEnd = 20
                    Case InStr(1, serviceId, "SP7", vbTextCompare) > 0
                        periodStart = 0
                        periodEnd = 20
                    Case InStr(1, serviceId, "ST9", vbTextCompare) > 0
                        periodStart = 0
                        periodEnd = 20
                    Case InStr(1, serviceId, "ST6", vbTextCompare) > 0
                        periodStart = 0
                        periodEnd = 13
                    Case InStr(1, serviceId, "Argo", vbTextCompare) > 0
                        periodStart = 0
                        periodEnd = 13
                    Case Else
                        periodStart = 0
                        periodEnd = 20
                End Select
            End If
        ' ГЛАВНОЕ: если марка = HAVAL - период БЕЗ 0, всегда начинается с 1
        ElseIf brand = "HAVAL" Then
