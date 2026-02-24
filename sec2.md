Option Explicit

Sub Process_IntegratedAP_To_2026()

    Dim wbInt As Workbook, wbUnix As Workbook, wb2026 As Workbook
    Dim wsInt As Worksheet, wsUnix As Worksheet, wsTarget As Worksheet
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim col As Long, r As Long
    Dim lastCol As Long, lastRow As Long
    Dim hostRow As Long
    
    Dim serverName As String
    Dim loginDate As String
    
    Dim accountName As String
    Dim sourceHost As String
    Dim loginText As String
    
    Dim startTime As String, endTime As String
    Dim key As String
    
    MsgBox "Select IntegratedAP workbook"
    Set wbInt = Workbooks.Open(Application.GetOpenFilename)
    
    MsgBox "Select UNIX Exclusion workbook"
    Set wbUnix = Workbooks.Open(Application.GetOpenFilename)
    
    MsgBox "Select 2026 workbook"
    Set wb2026 = Workbooks.Open(Application.GetOpenFilename)
    
    Set wsInt = wbInt.Worksheets("result_last")
    Set wsUnix = wbUnix.Worksheets("UNIX系ダイレクトログイン除外アカウント一覧")
    Set wsTarget = wb2026.Worksheets("ダイレクトログイン_202601")
    
    lastCol = wsInt.Cells(1, wsInt.Columns.Count).End(xlToLeft).Column
    
    For col = 2 To lastCol
        
        If wsInt.Cells(1, col).Value <> "" Then
            
            serverName = Trim(CStr(wsInt.Cells(2, col).Value))
            hostRow = 0
            
            For r = 1 To 50
                If InStr(wsInt.Cells(r, col).Value, "Host=") > 0 Then
                    hostRow = r
                    loginDate = SafeExtractDate(wsInt.Cells(r, col).Value)
                    Exit For
                End If
            Next r
            
            If hostRow = 0 Then GoTo NextColumn
            
            lastRow = wsInt.Cells(wsInt.Rows.Count, col).End(xlUp).Row
            
            For r = hostRow + 1 To lastRow
                
                If wsInt.Cells(r, col).DisplayFormat.Interior.Color = RGB(255, 0, 0) Then
                    
                    accountName = SafeLeftWord(wsInt.Cells(r, col).Value)
                    sourceHost = NormalizeHost(CStr(wsInt.Cells(r, col + 1).Value))
                    loginText = CStr(wsInt.Cells(r, col + 2).Value)
                    
                    If accountName = "" Or sourceHost = "" Then GoTo SkipRow
                    If IsExcluded(wsUnix, accountName, sourceHost) Then GoTo SkipRow
                    
                    ExtractTimeSafe loginText, startTime, endTime
                    
                    key = loginDate & "|" & serverName & "|" & accountName & "|" & sourceHost
                    
                    If Not dict.exists(key) Then
                        dict.Add key, Array(1, startTime, endTime)
                    Else
                        Dim arr As Variant
                        arr = dict(key)
                        
                        arr(0) = CLng(arr(0)) + 1
                        
                        If startTime <> "" Then
                            If arr(1) = "" Or startTime < arr(1) Then arr(1) = startTime
                        End If
                        
                        If endTime <> "" Then
                            If arr(2) = "" Or endTime > arr(2) Then arr(2) = endTime
                        End If
                        
                        dict(key) = arr
                    End If
                    
SkipRow:
                End If
            Next r
        End If
        
NextColumn:
    Next col
    
    ' ================= APPEND ONLY =================
    
    Dim insertRow As Long
    insertRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row + 1
    
    Dim parts() As String
    Dim resolvedIP As String
    Dim itemKey As Variant
    Dim dataArr As Variant
    
    For Each itemKey In dict.Keys
        
        parts = Split(CStr(itemKey), "|")
        dataArr = dict(itemKey)
        
        resolvedIP = GetIP(parts(3))
        
        wsTarget.Cells(insertRow, 2).Value = parts(0)
        wsTarget.Cells(insertRow, 4).Value = "GIB"
        wsTarget.Cells(insertRow, 5).Value = "Integrated AP"
        wsTarget.Cells(insertRow, 6).Value = parts(1)
        wsTarget.Cells(insertRow, 7).Value = GetDR(parts(1))
        wsTarget.Cells(insertRow, 8).Value = "Linux"
        wsTarget.Cells(insertRow, 9).Value = parts(2)
        wsTarget.Cells(insertRow, 10).Value = parts(3)
        wsTarget.Cells(insertRow, 11).Value = resolvedIP
        wsTarget.Cells(insertRow, 12).Value = CLng(dataArr(0))
        wsTarget.Cells(insertRow, 13).Value = dataArr(1) & " - " & dataArr(2)
        
        insertRow = insertRow + 1
        
    Next itemKey
    
    MsgBox "Process Completed Successfully!", vbInformation

End Sub

' ================= SAFE HELPERS =================

Function SafeExtractDate(textLine As String) As String
    Dim p As Long
    p = InStr(textLine, "Date=")
    If p = 0 Then
        SafeExtractDate = ""
        Exit Function
    End If
    
    SafeExtractDate = Trim(Replace(Mid(textLine, p + 5), """", ""))
End Function

Function SafeLeftWord(txt As String) As String
    If Trim(txt) = "" Then
        SafeLeftWord = ""
    Else
        SafeLeftWord = Trim(Split(Trim(txt), " ")(0))
    End If
End Function

Sub ExtractTimeSafe(fullText As String, ByRef startTime As String, ByRef endTime As String)

    Dim dashPos As Long
    dashPos = InStr(fullText, "-")
    
    If dashPos > 6 Then
        startTime = Trim(Mid(fullText, dashPos - 6, 5))
        endTime = Trim(Mid(fullText, dashPos + 2, 5))
    Else
        startTime = ""
        endTime = ""
    End If

End Sub

Function NormalizeHost(h As String) As String
    h = LCase(Trim(h))
    h = Replace(h, " ", "")
    h = Replace(h, "|", "l")
    If Left(h, 1) = "." Then h = Mid(h, 2)
    NormalizeHost = h
End Function

Function GetDR(serverName As String) As String
    Select Case LCase(serverName)
        Case "vrjpn40082", "vrjpn40083"
            GetDR = "Prod"
        Case "vrjpn40084"
            GetDR = "DR"
        Case Else
            GetDR = "Prod"
    End Select
End Function

Function IsExcluded(ws As Worksheet, acc As String, srv As String) As Boolean

    Dim r As Long
    
    For r = 4 To ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
        If LCase(CStr(ws.Cells(r, 3).Value)) = LCase(acc) Then
            If InStr(LCase(CStr(ws.Cells(r, 6).Value)), LCase(srv)) > 0 Then
                IsExcluded = True
                Exit Function
            End If
        End If
    Next r
    
    IsExcluded = False

End Function

Function GetIP(hostName As String) As String

    Dim sh As Object, exec As Object, output As String
    
    On Error Resume Next
    
    Set sh = CreateObject("WScript.Shell")
    Set exec = sh.Exec("powershell -Command ""Resolve-DnsName -Name '" & hostName & "' -ErrorAction SilentlyContinue | Where-Object {$_.IPAddress} | Select-Object -ExpandProperty IPAddress""")
    
    output = exec.StdOut.ReadAll
    
    If Trim(output) <> "" Then
        GetIP = Trim(Split(output, vbCrLf)(0))
    Else
        GetIP = ""
    End If

End Function
