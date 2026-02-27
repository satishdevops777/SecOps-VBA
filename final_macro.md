#Final_Macro
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
