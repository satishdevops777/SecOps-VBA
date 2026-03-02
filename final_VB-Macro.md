## PHASE-2 UNIX EXCLUSIONS
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
    Dim r As Long, uRow As Long
    
    Dim business As String
    Dim server As String
    Dim username As String
    
    Dim normalizedBaseHost As String
    Dim unixCellText As String
    
    Dim skipHost As Boolean
    Dim foundUser As Boolean

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
        
        normalizedBaseHost = NormalizeHost(wsPhase1.Cells(r, 4).Value)
        
        skipHost = False
        foundUser = False
        
        '---------------------------------------
        ' FULL SHEET SCAN FOR SAME USERNAME
        '---------------------------------------
        For uRow = 4 To unixLastRow
            
            If LCase(Trim(wsUnix.Cells(uRow, 3).Value)) = LCase(username) Then
                
                foundUser = True
                
                unixCellText = LCase(Trim(wsUnix.Cells(uRow, 6).Value))
                
                If unixCellText <> "" Then
                    
                    ' Remove domain from unix cell text also
                    unixCellText = Replace(unixCellText, ".prudential.com", "")
                    
                    ' Check if hostname exists anywhere in cell text
                    If InStr(unixCellText, normalizedBaseHost) > 0 Then
                        
                        skipHost = True
                        
                        WriteUnixLog wsLog, _
                            "User: " & username & _
                            " | Host: " & normalizedBaseHost & _
                            " | Status: Known Login (Excluded)"
                        
                        Exit For
                        
                    End If
                    
                End If
                
            End If
            
        Next uRow
        
        '---------------------------------------
        ' If user not found
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
        If skipHost = False Then
            
            wsOutput.Cells(outputRow, 1).Value = business
            wsOutput.Cells(outputRow, 2).Value = server
            wsOutput.Cells(outputRow, 3).Value = username
            wsOutput.Cells(outputRow, 4).Value = wsPhase1.Cells(r, 4).Value
            wsOutput.Cells(outputRow, 5).Value = wsPhase1.Cells(r, 5).Value
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


'=================================================
' Host Normalization Function
'=================================================
Function NormalizeHost(hostValue As String) As String

    Dim domainSuffix As String
    domainSuffix = ".prudential.com"
    
    hostValue = LCase(Trim(hostValue))
    
    If Right(hostValue, Len(domainSuffix)) = domainSuffix Then
        hostValue = Left(hostValue, Len(hostValue) - Len(domainSuffix))
    End If
    
    NormalizeHost = hostValue

End Function


'=================================================
' Logging Function
'=================================================
Sub WriteUnixLog(ws As Worksheet, msg As String)

    Dim nextRow As Long
    
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = msg

End Sub
```
