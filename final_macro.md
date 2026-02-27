#Final_Macro

## Phase-1 EXTRACTING VALUES
---
```vba
Option Explicit

Sub Extract_Suspicious_Summary()

    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsOutput As Worksheet
    
    Dim sourcePath As String
    Dim lastCol As Long
    Dim col As Long
    Dim r As Long
    
    Dim business As String
    Dim server As String
    Dim expectedCount As Long
    Dim redCount As Long
    
    Dim logLine As String
    Dim cleanLine As String
    Dim arr() As String
    
    Dim username As String
    Dim hostname As String
    Dim loginTime As String
    Dim logoutTime As String
    
    Dim firstLogin As String
    Dim lastLogout As String
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ResetLog
    WriteLog "Process Started"
    
    '=============================
    ' Open Source Workbook
    '=============================
    sourcePath = ThisWorkbook.Path & "\IntegratedAP_01_23Unixログチェックマクロ_ver71_20260215.xlsm"
    
    Set wbSource = Workbooks.Open(sourcePath, ReadOnly:=True)
    Set wsSource = wbSource.Sheets("result_last")
    
    WriteLog "Source Opened"
    
    '=============================
    ' Prepare Output Sheet
    '=============================
    Set wsOutput = PrepareOutput
    
    'Find last used column in row 1
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    
    '=============================
    ' Loop Through Columns (B onwards)
    '=============================
    For col = 2 To lastCol
        
        If Trim(wsSource.Cells(1, col).Value) = "" Then Exit For
        
        business = Trim(wsSource.Cells(1, col).Value)
        server = Trim(wsSource.Cells(2, col).Value)
        expectedCount = Val(wsSource.Cells(12, col).Value)
        
        WriteLog "Processing Column " & col & " | Server: " & server
        
        redCount = 0
        firstLogin = ""
        lastLogout = ""
        
        '=============================
        ' Scan Rows From 20 Downwards
        '=============================
        r = 21
        
        Do While Trim(wsSource.Cells(r, col).Value) <> ""
            
            'Debug: Log color info
            WriteLog "Row " & r & _
                     " | Interior: " & wsSource.Cells(r, col).Interior.Color
            
            'Check PURE RED
            If wsSource.Cells(r, col).Interior.Color = 255 Then
                
                redCount = redCount + 1
                
                logLine = wsSource.Cells(r, col).Value
                cleanLine = Application.WorksheetFunction.Trim(logLine)
                
                arr = Split(cleanLine, " ")
                
                'Safety check
                If UBound(arr) >= 8 Then
                    
                    username = arr(0)
                    hostname = arr(2)
                    loginTime = arr(6)
                    logoutTime = arr(8)
                    
                    'Find EARLIEST login
                    If firstLogin = "" Then
                        firstLogin = loginTime
                    Else
                        If TimeValue(loginTime) < TimeValue(firstLogin) Then
                            firstLogin = loginTime
                        End If
                    End If
                    
                    'Find LATEST logout
                    If lastLogout = "" Then
                        lastLogout = logoutTime
                    Else
                        If TimeValue(logoutTime) > TimeValue(lastLogout) Then
                            lastLogout = logoutTime
                        End If
                    End If
                    
                End If
                
            End If
            
            r = r + 1
            
        Loop
        
        '=============================
        ' Validate Count
        '=============================
        If redCount <> expectedCount Then
            WriteLog "Count Mismatch in column " & col & _
                     " Expected: " & expectedCount & _
                     " Found: " & redCount
        Else
            WriteLog "Count Matched: " & redCount
        End If
        
        '=============================
        ' Write Summary Row
        '=============================
        If redCount > 0 Then
            WriteToOutput wsOutput, business, server, _
                          username, hostname, _
                          redCount, firstLogin, lastLogout
        End If
        
    Next col
    
    wbSource.Close False
    
    WriteLog "Process Completed"
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "Summary Extracted Successfully", vbInformation

End Sub


Sub ResetLog()

    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Log")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Log"
    Else
        ws.Cells.Clear
    End If
    
    ws.Range("A1").Value = "Timestamp"
    ws.Range("B1").Value = "Message"

End Sub


Sub WriteLog(msg As String)

    Dim ws As Worksheet
    Dim nextRow As Long
    
    Set ws = ThisWorkbook.Sheets("Log")
    
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = msg

End Sub

Function PrepareOutput() As Worksheet

    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Output")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Output"
    Else
        ws.Cells.Clear
    End If
    
    ws.Range("A1").Value = "Business"
    ws.Range("B1").Value = "Server"
    ws.Range("C1").Value = "Username"
    ws.Range("D1").Value = "Hostname"
    ws.Range("E1").Value = "Count"
    ws.Range("F1").Value = "Login Time"
    ws.Range("G1").Value = "Logout Time"
    
    Set PrepareOutput = ws

End Function

Sub WriteToOutput(ws As Worksheet, b As String, s As String, _
                  u As String, h As String, c As Long, _
                  l As String, o As String)

    Dim nextRow As Long
    
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ws.Cells(nextRow, 1).Value = b
    ws.Cells(nextRow, 2).Value = s
    ws.Cells(nextRow, 3).Value = u
    ws.Cells(nextRow, 4).Value = h
    ws.Cells(nextRow, 5).Value = c
    ws.Cells(nextRow, 6).Value = l
    ws.Cells(nextRow, 7).Value = o

End Sub

```
---
## PHASE-2 UNIX EXCLUSIONS

### Final-Phase2
```vba

Option Explicit

Sub Phase2_CheckUnixExclusion()

    '=============================
    ' >>> EDIT THESE TWO VALUES <<<
    '=============================
    Dim unixWorkbookName As String
    Dim unixSheetName As String
    
    unixWorkbookName = "UNIX.xlsx"
    unixSheetName = "UNIX系ダイレクトログイン除外アカウント一覧"
    '=============================

    Dim wsPhase1 As Worksheet
    Dim wsUnix As Worksheet
    Dim wsOutput As Worksheet
    Dim wsLog As Worksheet
    Dim wbUnix As Workbook
    
    Dim lastRow As Long
    Dim unixLastRow As Long
    Dim r As Long, uRow As Long, checkRow As Long
    
    Dim business As String
    Dim server As String
    Dim username As String
    
    Dim fullHost As String
    Dim normalizedBaseHost As String
    Dim normalizedExHost As String
    
    Dim foundUser As Boolean
    Dim skipHost As Boolean

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    '=============================
    ' Prepare Log Sheet
    '=============================
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("Unix_Exclusion_Logs")
    On Error GoTo 0
    
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add
        wsLog.Name = "Unix_Exclusion_Logs"
    Else
        wsLog.Cells.Clear
    End If
    
    wsLog.Range("A1:B1").Value = Array("Timestamp", "Message")
    WriteUnixLog wsLog, "Phase 2 Started"
    
    Set wsPhase1 = ThisWorkbook.Sheets("Output")
    
    '=============================
    ' Prepare Output Sheet
    '=============================
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Sheets("Unix_Exclusion_Output")
    On Error GoTo 0
    
    If wsOutput Is Nothing Then
        Set wsOutput = ThisWorkbook.Sheets.Add
        wsOutput.Name = "Unix_Exclusion_Output"
    Else
        wsOutput.Cells.Clear
    End If
    
    wsOutput.Range("A1:G1").Value = Array( _
        "Business", "Server", "Username", _
        "Hostname", "Count", "Login", "Logout")
    
    WriteUnixLog wsLog, "Opening Unix Workbook: " & unixWorkbookName
    
    Set wbUnix = Workbooks.Open(ThisWorkbook.Path & "\" & unixWorkbookName, ReadOnly:=True)
    Set wsUnix = wbUnix.Sheets(unixSheetName)
    
    lastRow = wsPhase1.Cells(wsPhase1.Rows.Count, 1).End(xlUp).Row
    unixLastRow = wsUnix.Cells(wsUnix.Rows.Count, 3).End(xlUp).Row
    
    Dim outputRow As Long
    outputRow = 2
    
    '=============================
    ' MAIN LOOP
    '=============================
    For r = 2 To lastRow
        
        business = wsPhase1.Cells(r, 1).Value
        server = wsPhase1.Cells(r, 2).Value
        username = Trim(wsPhase1.Cells(r, 3).Value)
        
        fullHost = LCase(Trim(wsPhase1.Cells(r, 4).Value))
        
        If InStr(fullHost, ".") > 0 Then
            normalizedBaseHost = Split(fullHost, ".")(0)
        Else
            normalizedBaseHost = fullHost
        End If
        
        foundUser = False
        skipHost = False
        
        '---------------------------------------
        ' Find user block in UNIX sheet
        '---------------------------------------
        For uRow = 4 To unixLastRow
            
            If LCase(Trim(wsUnix.Cells(uRow, 3).Value)) = LCase(username) Then
                
                foundUser = True
                checkRow = uRow
                
                Do While checkRow <= unixLastRow
                    
                    If checkRow > uRow Then
                        If Trim(wsUnix.Cells(checkRow, 3).Value) <> "" Then Exit Do
                    End If
                    
                    normalizedExHost = LCase(Trim(wsUnix.Cells(checkRow, 6).Value))
                    
                    If normalizedExHost <> "" Then
                        
                        If InStr(normalizedExHost, ".") > 0 Then
                            normalizedExHost = Split(normalizedExHost, ".")(0)
                        End If
                        
                        If normalizedExHost = normalizedBaseHost Then
                            skipHost = True
                            WriteUnixLog wsLog, _
                                "User: " & username & _
                                " | Host: " & normalizedBaseHost & _
                                " | Status: Known Login (Excluded)"
                            Exit Do
                        End If
                        
                    End If
                    
                    checkRow = checkRow + 1
                    
                Loop
                
                Exit For
                
            End If
            
        Next uRow
        
        '---------------------------------------
        ' If user not found in UNIX sheet
        '---------------------------------------
        If foundUser = False Then
            WriteUnixLog wsLog, _
                "User: " & username & _
                " | Host: " & normalizedBaseHost & _
                " | Status: User Not Found in UNIX Sheet"
        End If
        
        '---------------------------------------
        ' Print only if NOT excluded
        '---------------------------------------
        If foundUser = True And skipHost = False Then
            
            wsOutput.Cells(outputRow, 1).Value = business
            wsOutput.Cells(outputRow, 2).Value = server
            wsOutput.Cells(outputRow, 3).Value = username
            wsOutput.Cells(outputRow, 4).Value = wsPhase1.Cells(r, 4).Value
            wsOutput.Cells(outputRow, 5).Value = wsPhase1.Cells(r, 5).Value
            
            ' EXACT COPY (NO CONVERSION)
            wsOutput.Cells(outputRow, 6).Value = wsPhase1.Cells(r, 6).Value
            wsOutput.Cells(outputRow, 7).Value = wsPhase1.Cells(r, 7).Value
            
            WriteUnixLog wsLog, _
                "User: " & username & _
                " | Host: " & normalizedBaseHost & _
                " | Status: Unknown Login (Printed)"
            
            outputRow = outputRow + 1
            
        End If
        
    Next r
    
    wbUnix.Close False
    
    WriteUnixLog wsLog, "Phase 2 Completed"
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

Sub WriteUnixLog(ws As Worksheet, msg As String)

    Dim nextRow As Long
    
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = msg

End Sub
```
---
```vba
Option Explicit

Sub Phase2_CheckUnixExclusion()

    '=============================
    ' >>> EDIT ONLY THESE TWO <<<
    '=============================
    Dim unixWorkbookName As String
    Dim unixSheetName As String
    
    unixWorkbookName = "UNIX.xlsx"
    unixSheetName = "UNIX系ダイレクトログイン除外アカウント一覧"
    '=============================

    Dim wsPhase1 As Worksheet
    Dim wsUnix As Worksheet
    Dim wsOutput As Worksheet
    Dim wbUnix As Workbook
    
    Dim lastRow As Long
    Dim unixLastRow As Long
    Dim r As Long, uRow As Long
    
    Dim business As String
    Dim server As String
    Dim username As String
    Dim fullHost As String
    Dim baseHost As String
    Dim exclusionText As String
    
    Dim foundUser As Boolean
    Dim skipHost As Boolean
    
    Dim printedCount As Long
    Dim skippedCount As Long
    Dim notFoundCount As Long
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ResetLog
    WriteLog "Phase 2 Started"
    
    'Phase 1 Output sheet
    Set wsPhase1 = ThisWorkbook.Sheets("Output")
    
    'Create / Clear Output Sheet
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Sheets("Unix_Exclusion_Output")
    On Error GoTo 0
    
    If wsOutput Is Nothing Then
        Set wsOutput = ThisWorkbook.Sheets.Add
        wsOutput.Name = "Unix_Exclusion_Output"
    Else
        wsOutput.Cells.Clear
    End If
    
    wsOutput.Range("A1:G1").Value = Array("Business", "Server", "Username", _
                                          "Hostname", "Count", "Login", "Logout")
    
    WriteLog "Opening Unix Workbook: " & unixWorkbookName
    
    Set wbUnix = Workbooks.Open(ThisWorkbook.Path & "\" & unixWorkbookName, ReadOnly:=True)
    Set wsUnix = wbUnix.Sheets(unixSheetName)
    
    lastRow = wsPhase1.Cells(wsPhase1.Rows.Count, 1).End(xlUp).Row
    unixLastRow = wsUnix.Cells(wsUnix.Rows.Count, 3).End(xlUp).Row
    
    Dim outputRow As Long
    outputRow = 2
    
    For r = 2 To lastRow
        
        business = wsPhase1.Cells(r, 1).Value
        server = wsPhase1.Cells(r, 2).Value
        username = Trim(wsPhase1.Cells(r, 3).Value)
        
        fullHost = LCase(Trim(wsPhase1.Cells(r, 4).Value))
        baseHost = Split(fullHost, ".")(0)
        
        foundUser = False
        skipHost = False
        
        For uRow = 4 To unixLastRow
            
            If LCase(Trim(wsUnix.Cells(uRow, 3).Value)) = LCase(username) Then
                
                foundUser = True
                exclusionText = LCase(wsUnix.Cells(uRow, 6).Value)
                
                If InStr(1, exclusionText, baseHost, vbTextCompare) > 0 Then
                    skipHost = True
                    skippedCount = skippedCount + 1
                    WriteLog "Skipped (Excluded): " & username & " | " & baseHost
                End If
                
                Exit For
                
            End If
            
        Next uRow
        
        If foundUser = False Then
            notFoundCount = notFoundCount + 1
            WriteLog "Username Not Found in Unix Sheet: " & username
        End If
        
        If foundUser = True And skipHost = False Then
            
            wsOutput.Cells(outputRow, 1).Value = business
            wsOutput.Cells(outputRow, 2).Value = server
            wsOutput.Cells(outputRow, 3).Value = username
            wsOutput.Cells(outputRow, 4).Value = wsPhase1.Cells(r, 4).Value
            wsOutput.Cells(outputRow, 5).Value = wsPhase1.Cells(r, 5).Value
            wsOutput.Cells(outputRow, 6).Value = wsPhase1.Cells(r, 6).Value
            wsOutput.Cells(outputRow, 7).Value = wsPhase1.Cells(r, 7).Value
            
            printedCount = printedCount + 1
            WriteLog "Printed: " & username & " | " & baseHost
            
            outputRow = outputRow + 1
            
        End If
        
    Next r
    
    wbUnix.Close False
    
    WriteLog "Phase 2 Completed"
    WriteLog "Total Printed: " & printedCount
    WriteLog "Total Skipped (Excluded): " & skippedCount
    WriteLog "Total Username Not Found: " & notFoundCount
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub
```


```vba
Option Explicit

Sub Phase2_CheckUnixExclusion()

    '=============================
    ' >>> EDIT THESE TWO VALUES <<<
    '=============================
    Dim unixWorkbookName As String
    Dim unixSheetName As String
    
    unixWorkbookName = "UNIX.xlsm"
    unixSheetName = "UNIX系ダイレクトログイン除外アカウント一覧"
    '=============================

    Dim wsPhase1 As Worksheet
    Dim wsUnix As Worksheet
    Dim wsOutput As Worksheet
    Dim wsLog As Worksheet
    Dim wbUnix As Workbook
    
    Dim lastRow As Long
    Dim unixLastRow As Long
    Dim r As Long, uRow As Long, checkRow As Long
    
    Dim business As String
    Dim server As String
    Dim username As String
    
    Dim fullHost As String
    Dim normalizedBaseHost As String
    Dim normalizedExHost As String
    
    Dim foundUser As Boolean
    Dim skipHost As Boolean
    Dim t1 As Date, t2 As Date
    Dim finalLogin As Date, finalLogout As Date
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    '=============================
    ' Prepare Log Sheet
    '=============================
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("Unix_Exclusion_Logs")
    On Error GoTo 0
    
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add
        wsLog.Name = "Unix_Exclusion_Logs"
    Else
        wsLog.Cells.Clear
    End If
    
    wsLog.Range("A1:B1").Value = Array("Timestamp", "Message")
    WriteUnixLog wsLog, "Phase 2 Started"
    
    Set wsPhase1 = ThisWorkbook.Sheets("Output")
    
    '=============================
    ' Prepare Output Sheet
    '=============================
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Sheets("Unix_Exclusion_Output")
    On Error GoTo 0
    
    If wsOutput Is Nothing Then
        Set wsOutput = ThisWorkbook.Sheets.Add
        wsOutput.Name = "Unix_Exclusion_Output"
    Else
        wsOutput.Cells.Clear
    End If
    
    wsOutput.Range("A1:G1").Value = Array("Business", "Server", "Username", _
                                          "Hostname", "Count", "Login", "Logout")
    
    WriteUnixLog wsLog, "Opening Unix Workbook: " & unixWorkbookName
    
    Set wbUnix = Workbooks.Open(ThisWorkbook.Path & "\" & unixWorkbookName, ReadOnly:=True)
    Set wsUnix = wbUnix.Sheets(unixSheetName)
    
    lastRow = wsPhase1.Cells(wsPhase1.Rows.Count, 1).End(xlUp).Row
    unixLastRow = wsUnix.Cells(wsUnix.Rows.Count, 3).End(xlUp).Row
    
    Dim outputRow As Long
    outputRow = 2
    
    '=============================
    ' MAIN LOOP
    '=============================
    For r = 2 To lastRow
        
        business = wsPhase1.Cells(r, 1).Value
        server = wsPhase1.Cells(r, 2).Value
        username = Trim(wsPhase1.Cells(r, 3).Value)
        
        fullHost = LCase(Trim(wsPhase1.Cells(r, 4).Value))
        
        If InStr(fullHost, ".") > 0 Then
            normalizedBaseHost = Split(fullHost, ".")(0)
        Else
            normalizedBaseHost = fullHost
        End If
        
        foundUser = False
        skipHost = False
        
        '---------------------------------------
        ' Find user block in UNIX sheet
        '---------------------------------------
        For uRow = 4 To unixLastRow
            
            If LCase(Trim(wsUnix.Cells(uRow, 3).Value)) = LCase(username) Then
                
                foundUser = True
                checkRow = uRow
                
                Do While checkRow <= unixLastRow
                    
                    If checkRow > uRow Then
                        If Trim(wsUnix.Cells(checkRow, 3).Value) <> "" Then Exit Do
                    End If
                    
                    normalizedExHost = LCase(Trim(wsUnix.Cells(checkRow, 6).Value))
                    
                    If normalizedExHost <> "" Then
                        
                        If InStr(normalizedExHost, ".") > 0 Then
                            normalizedExHost = Split(normalizedExHost, ".")(0)
                        End If
                        
                        If normalizedExHost = normalizedBaseHost Then
                            skipHost = True
                            WriteUnixLog wsLog, "User: " & username & _
                                                 " | Host: " & normalizedBaseHost & _
                                                 " | Status: Known Login (Excluded)"
                            Exit Do
                        End If
                        
                    End If
                    
                    checkRow = checkRow + 1
                    
                Loop
                
                Exit For
                
            End If
            
        Next uRow
        
        '---------------------------------------
        ' If user not found in UNIX sheet
        '---------------------------------------
        If foundUser = False Then
            WriteUnixLog wsLog, "User: " & username & _
                                 " | Host: " & normalizedBaseHost & _
                                 " | Status: User Not Found in UNIX Sheet"
        End If
        
        '---------------------------------------
        ' Print only if NOT excluded
        '---------------------------------------
        If foundUser = True And skipHost = False Then
            
            wsOutput.Cells(outputRow, 1).Value = business
            wsOutput.Cells(outputRow, 2).Value = server
            wsOutput.Cells(outputRow, 3).Value = username
            wsOutput.Cells(outputRow, 4).Value = wsPhase1.Cells(r, 4).Value
            wsOutput.Cells(outputRow, 5).Value = wsPhase1.Cells(r, 5).Value  
            t1 = TimeValue(wsPhase1.Cells(r, 6).Value)
            t2 = TimeValue(wsPhase1.Cells(r, 7).Value)
            If t1 <= t2 Then
                finalLogin = t1
                finalLogout = t2
            Else
                finalLogin = t2
                finalLogout = t1
            End If
            wsOutput.Cells(outputRow, 6).Value = Format(finalLogin, "hh:mm")
            wsOutput.Cells(outputRow, 7).Value = Format(finalLogout, "hh:mm")
            
            WriteUnixLog wsLog, "User: " & username & _
                                 " | Host: " & normalizedBaseHost & _
                                 " | Status: Unknown Login (Printed)"
            
            outputRow = outputRow + 1
            
        End If
        
    Next r
    
    wbUnix.Close False
    
    WriteUnixLog wsLog, "Phase 2 Completed"
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

Sub WriteUnixLog(ws As Worksheet, msg As String)

    Dim nextRow As Long
    
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = msg

End Sub
```

## PHASE-3 
---
```vba
Option Explicit

Sub Phase3_UpdateTracking()

    Dim targetWb As Workbook
    Dim wsTarget As Worksheet
    Dim wsOutput As Worksheet
    
    Dim targetPath As String
    Dim targetFile As String
    Dim targetSheet As String
    
    Dim lastOutputRow As Long
    Dim lastTargetRow As Long
    Dim nextNo As Long
    Dim r As Long
    
    Dim business As String
    Dim personInCharge As String
    Dim buValue As String
    Dim systemName As String
    
    Dim server As String
    Dim drProd As String
    Dim fqdn As String
    Dim ipAddress As String
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Phase3_ResetLog
    Phase3_WriteLog "Phase 3 Process Started"
    
    targetFile = "20260117～発生_ダイレクトログイン発生分.xlsx"
    targetSheet = "ダイレクトログイン_202601"
    targetPath = ThisWorkbook.Path & "\" & targetFile
    
    On Error Resume Next
    Set targetWb = Workbooks(targetFile)
    On Error GoTo 0
    
    If targetWb Is Nothing Then
        Set targetWb = Workbooks.Open(targetPath)
        Phase3_WriteLog "Target Workbook Opened"
    End If
    
    Set wsTarget = targetWb.Sheets(targetSheet)
    Set wsOutput = ThisWorkbook.Sheets("Unix_Exclusion_Output")
    
    lastOutputRow = wsOutput.Cells(wsOutput.Rows.Count, 1).End(xlUp).Row
    lastTargetRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row
    
    nextNo = wsTarget.Cells(lastTargetRow, 1).Value + 1
    
    For r = 2 To lastOutputRow
        
        business = Trim(wsOutput.Cells(r, 1).Value)
        server = Trim(wsOutput.Cells(r, 2).Value)
        fqdn = Trim(wsOutput.Cells(r, 4).Value)
        
        Phase3_WriteLog "Processing User: " & wsOutput.Cells(r, 3).Value & _
                        " | Host: " & fqdn
        
        '========================
        ' BUSINESS MAPPING
        '========================
        If LCase(business) = "integratedap_01" Then
            personInCharge = "プラギャ"
            buValue = "GIB"
            systemName = "Integrated AP"
        Else
            personInCharge = ""
            buValue = ""
            systemName = business
        End If
        
        '========================
        ' DR / PROD LOGIC
        '========================
        If LCase(server) = "vrjpn40084" Then
            drProd = "DR"
        ElseIf LCase(server) = "vrjpn40082" Or _
               LCase(server) = "vrjpn40083" Then
            drProd = "Prod"
        Else
            drProd = ""
        End If
        
        '========================
        ' RESOLVE IP
        '========================
        ipAddress = GetIPAddressFromFQDN(fqdn)
        
        If ipAddress = "" Then
            Phase3_WriteLog "IP NOT FOUND for " & fqdn
        Else
            Phase3_WriteLog "IP Resolved: " & ipAddress
        End If
        
        lastTargetRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row + 1
        
        wsTarget.Cells(lastTargetRow, 1).Value = nextNo
        wsTarget.Cells(lastTargetRow, 2).Value = Date - 1
        wsTarget.Cells(lastTargetRow, 3).Value = personInCharge
        wsTarget.Cells(lastTargetRow, 4).Value = buValue
        wsTarget.Cells(lastTargetRow, 5).Value = systemName
        wsTarget.Cells(lastTargetRow, 6).Value = server
        wsTarget.Cells(lastTargetRow, 7).Value = drProd
        wsTarget.Cells(lastTargetRow, 8).Value = "Linux"
        wsTarget.Cells(lastTargetRow, 9).Value = wsOutput.Cells(r, 3).Value
        wsTarget.Cells(lastTargetRow, 10).Value = fqdn
        wsTarget.Cells(lastTargetRow, 11).Value = ipAddress
        wsTarget.Cells(lastTargetRow, 12).Value = wsOutput.Cells(r, 5).Value
        wsTarget.Cells(lastTargetRow, 13).Value = _
            wsOutput.Cells(r, 6).Value & " - " & wsOutput.Cells(r, 7).Value
        
        Phase3_WriteLog "Row Added with NO: " & nextNo
        
        nextNo = nextNo + 1
        
    Next r
    
    targetWb.Save
    Phase3_WriteLog "Workbook Saved"
    Phase3_WriteLog "Phase 3 Completed Successfully"
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub


Sub Phase3_ResetLog()

    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Phase3_Log")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Phase3_Log"
    Else
        ws.Cells.Clear
    End If
    
    ws.Range("A1").Value = "Timestamp"
    ws.Range("B1").Value = "Message"

End Sub

Sub Phase3_WriteLog(msg As String)

    Dim ws As Worksheet
    Dim nextRow As Long
    
    Set ws = ThisWorkbook.Sheets("Phase3_Log")
    
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = msg

End Sub


Function GetIPAddressFromFQDN(fqdn As String) As String

    Dim shell As Object
    Dim exec As Object
    Dim line As String
    
    On Error Resume Next
    
    Set shell = CreateObject("WScript.Shell")
    Set exec = shell.Exec("powershell -command ""Resolve-DnsName " & fqdn & " | Select-Object -ExpandProperty IPAddress""")
    
    Do While Not exec.StdOut.AtEndOfStream
        line = exec.StdOut.ReadLine
        If line <> "" Then
            GetIPAddressFromFQDN = line
            Exit Function
        End If
    Loop
    
    GetIPAddressFromFQDN = ""

End Function

```
