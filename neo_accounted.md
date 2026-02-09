

## 1️⃣ Param_Sheet_last_su (Excel control sheet)
### 📥 Input source: 
- Sheet: Param_Sheet_last_su

### 📌 Code where it is read (NEO_Accounted)

### 🧠 How it controls execution


| Column     | Effect in code                             |
| ---------- | ------------------------------------------ |
| A (System) | Used in output & summary                   |
| B (Y/N)    | Skips system if not `Y`                    |
| D + E      | Build file paths                           |
| H          | Chooses `Sec_Linux_Main` or `Sec_AIX_Main` |
| J          | Decides which syslog file to open          |


## 2️⃣ Syslog file (Linux / AIX)
### 📥 Input source
```
In_Path1 + In_Path2 + In_File_name_syslog
```


### 📌 Where it is opened
```

Linux:
Workbooks.Open Filename:=In_Path1 & In_Path2 & In_File_name_syslog, ReadOnly:=True
' #comment: Opens Linux syslog file


AIX:

Workbooks.Open Filename:=In_Path1 & In_Path2 & In_File_name_syslog, ReadOnly:=True
' #comment: Opens AIX syslog file
```

### 📌 What is extracted from syslog
```
 1. Timestamp

Edit_Date_Log = Left(ActiveCell.Offset(i, 0), 16)
' #comment: Extracts datetime from log line

Used to generate:

Edit_Date_Log_mmdd = Format(Edit_Date_Log, "mmdd")
Edit_Date_Log_hhmmss = Format(Edit_Date_Log, "hh:mm:ss")
```

### 2. User information    
```
Wk_Column = InStr(ActiveCell, "USER=")
Wk_User = Tmp_Str1(0)

Wk_Column = InStr(ActiveCell, "sudo:")
Wk_User = Tmp_Str1(0)
```

### 📌 Extracts:
- user before su
- user after su

### 3. PTS / TTY

```
Wk_Column = InStr(ActiveCell, "pts/")
Wk_Pts1 = "pts/" & Tmp_Str1(0)


Fallback:

If InStr(ActiveCell, "TTY=unknown") Then
    Wk_Pts1 = "unknown"
End If
```

### 4. Noise / exclusion detection

```
If InStr(ActiveCell, "closed") Or InStr(ActiveCell, "by uid") Then
    ActiveCell.Offset(i, 11) = 1
End If


If InStr(ActiveCell, "pam_vas: Authentication ignored") Then
    ActiveCell.Offset(i, 11) = 1
End If
```

### 📌 These lines do not become output records.





## 3️⃣ Script log files (correlation input)
### 📥 Input source
```
Dir(In_Path1 & In_Path2 & "*")
```

### 📌 Where they are read
```
Sub Sec_Script_Spool()
```

### 📌 What is extracted
- From filename:
```
Tmp_Script = Split(StrFileName, "_")

```
- Stored into:
```
Script_TBI_Account(Idx)
Script_TBI_St_Time(Idx)
Script_TBI_End_Time(Idx)
Script_TBI_Kanri_No(Idx)
Script_TBI_Script_File_Name(Idx)
```

### 📌 Purpose:
- Match syslog su time with script execution window
- Resolve Kanri No
- Resolve script file name



## 4️⃣ OUTPUT: NEO_Work_Sheet (main audit result)
### 📤 Where written

- Throughout Linux/AIX parsers:
```
ThisWorkbook.Sheets("NEO_Work_Sheet").Cells(Loop_Neo, 1) = "NEO_SU"
ThisWorkbook.Sheets("NEO_Work_Sheet").Cells(Loop_Neo, 2) = Edit_Date_Log_mmdd
ThisWorkbook.Sheets("NEO_Work_Sheet").Cells(Loop_Neo, 3) = Edit_Date_Log_hhmmss
ThisWorkbook.Sheets("NEO_Work_Sheet").Cells(Loop_Neo, 5) = Wk_Pts1
ThisWorkbook.Sheets("NEO_Work_Sheet").Cells(Loop_Neo, 6) = Wk_User & " → " & Wk_After_su
ThisWorkbook.Sheets("NEO_Work_Sheet").Cells(Loop_Neo, 8) = Kanri_No
ThisWorkbook.Sheets("NEO_Work_Sheet").Cells(Loop_Neo, 9) = Script_File_Name
```

### 📌 This is the canonical audit table.


## 5️⃣ OUTPUT: CSV file
### 📤 File produced
```
NEOsu_<MMDD>_accounted.csv
```

### 📌 Code
```
csvFile = In_Path1 & In_Path2 & "NEOsu_" & Edit_Date_Prm & "_accounted.csv"
Open csvFile For Output As #1
```

### Written from:
```
Sheets("NEO_Work_Csv")
```

###  📌 This is the compliance / SOC deliverable.


## 6️⃣ OUTPUT: SU_Count (summary)
### 📤 Purpose
- Per-system aggregation
- Counts by category
- Visual reporting

### 📌 Source data
- Built from NEO_Work_Sheet
- Uses system color from Param_Sheet_last_su





```
In_System = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 1)
' #comment: Reads system/host name

Chk_Ctl = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 2)
' #comment: Reads Y/N control flag

In_Path1 = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 4)
' #comment: Base directory for logs

In_Path2 = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 5)
' #comment: Subdirectory for logs

Chk_OS = UCase(ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 8))
' #comment: OS type (LINUX / AIX)

In_File_name_syslog = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 10)
' #comment: Syslog filename
```






```vba
Option Explicit
' #comment: Forces explicit variable declaration to avoid hidden bugs

' =====================================================
' Global variables shared across multiple procedures
' =====================================================

Public Now_ymd As String
' #comment: Stores current date/time for comparison and validation

Public Idx As Byte
' #comment: Generic index used for script-log arrays

Public Script_TBI_Account(50)
' #comment: Stores account names parsed from script log filenames

Public Script_TBI_St_Time(50)
' #comment: Stores script start timestamps

Public Script_TBI_End_Time(50)
' #comment: Stores script end timestamps

Public Script_TBI_Kanri_No(50)
' #comment: Stores management (Kanri) numbers per script log

Public Script_TBI_Script_File_Name(50)
' #comment: Stores script log file names

Public Edit_Date_Log As Date
' #comment: Full datetime parsed from syslog line

Public In_Path1 As String
' #comment: Base directory path for logs

Public In_Path2 As String
' #comment: Sub-directory path for logs

Public Wk_User As String
' #comment: Working variable holding detected user

Public Kanri_No As String
' #comment: Management number associated with SU operation

Public Script_File_Name As String
' #comment: Script log file name linked to SU event

Public In_System As String
' #comment: System/host name being processed

Public In_System_Color As Long
' #comment: Cell color of system name (used later in summary)

Public In_File_name_syslog As String
' #comment: Syslog file name to open

Public Edit_Date_Prm As String
' #comment: Date (MMDD) extracted from syslog filename

Public Edit_Date_Log_mmdd As String
' #comment: MMDD extracted from log timestamp

Public Edit_Date_Log_hhmmss As String
' #comment: HH:MM:SS extracted from log timestamp


' =====================================================
' PART 1: Main entry point
' =====================================================
Sub NEO_Accounted()

    Dim loop_i As Integer
    ' #comment: Row index for parameter sheet

    Dim Loop_Neo As Integer
    ' #comment: Output row counter for NEO_Work_Sheet

    Dim Wk_Activecell As String
    ' #comment: Used to detect end of parameter rows

    Dim Chk_Ctl As String
    ' #comment: Enable/disable flag per system

    Dim Chk_OS As String
    ' #comment: OS type (LINUX / AIX)

    Dim SU_Count_Idx As Byte
    ' #comment: Index for SU aggregation arrays

    Now_ymd = Now
    ' #comment: Capture current date/time

    Wk_Activecell = "9999"
    ' #comment: Dummy value to enter Do loop

    loop_i = 2
    ' #comment: Start from row 2 (skip header)

    SU_Count_Idx = 0
    ' #comment: Reset SU counter index

    Application.ScreenUpdating = False
    ' #comment: Disable Excel screen refresh for speed

    Do Until Wk_Activecell = ""

        In_System = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 1)
        ' #comment: Read system name

        In_System_Color = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 1).Interior.Color
        ' #comment: Save background color of system cell

        Chk_Ctl = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 2)
        ' #comment: Read execution control flag

        Chk_OS = UCase(ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 8))
        ' #comment: Read OS type and normalize

        If Chk_Ctl = "Y" Or Chk_Ctl = "y" Then
            ' #comment: Process only enabled rows

            In_Path1 = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 4)
            ' #comment: Read base path

            In_Path2 = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 5)
            ' #comment: Read sub path

            In_File_name_syslog = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 10)
            ' #comment: Read syslog filename

            Edit_Date_Prm = Left(Right(In_File_name_syslog, 8), 4)
            ' #comment: Extract MMDD from filename

            Erase Pts_Map_Tbl
            ' #comment: Clear PTS mapping table

            Call Sec_Script_Spool
            ' #comment: Preload script-log metadata

            Select Case Chk_OS
                Case "LINUX"
                    Call Sec_Linux_Main
                    ' #comment: Execute Linux syslog parser

                Case "AIX"
                    Call Sec_AIX_Main
                    ' #comment: Execute AIX syslog parser

                Case Else
                    MsgBox "LINUX / AIX parameter error"
                    ' #comment: Invalid OS parameter
            End Select
        End If

        loop_i = loop_i + 1
        ' #comment: Move to next parameter row

        Wk_Activecell = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 1)
        ' #comment: Check if next system exists

    Loop

    Application.ScreenUpdating = True
    ' #comment: Restore screen updates

    MsgBox "fin"
    ' #comment: End notification

End Sub


' =====================================================
' PART 2: Linux syslog parsing
' =====================================================
Sub Sec_Linux_Main()

    Dim Loop_Neo As Integer
    ' #comment: Output row counter

    Dim Wk_Str1 As String
    ' #comment: Temporary string buffer

    Dim Wk_Before_su As String
    ' #comment: User before su

    Dim Wk_After_su As String
    ' #comment: User after su

    Dim Wk_Column As Byte
    ' #comment: Column index for string parsing

    Dim Tmp_Str1 As Variant
    ' #comment: Array for Split() results

    Dim Wk_Pts1 As String
    ' #comment: PTS value (e.g., pts/3)

    Dim Pts_Idx As Byte
    ' #comment: Index for PTS map table

    Dim i As Long
    ' #comment: Row loop index

    Dim Date_i As Long
    ' #comment: Counter for date validation

    Dim Wk_i As Long
    ' #comment: Last row index

    Dim Wk_Hit_Flg As Byte
    ' #comment: Flag indicating match found

    Dim Write_Ctl As Byte
    ' #comment: Controls whether record is written

    Erase Pts_Map_Tbl
    ' #comment: Reset PTS mapping table

    Loop_Neo = 1
    ' #comment: Initialize output row

    ThisWorkbook.Sheets("NEO_Work_Sheet").Cells.Clear
    ' #comment: Clear previous Linux results

    ThisWorkbook.Sheets("Neo_Work_Csv").Cells.Clear
    ' #comment: Clear CSV work area

    Workbooks.Open Filename:=In_Path1 & In_Path2 & In_File_name_syslog, ReadOnly:=True
    ' #comment: Open syslog file as workbook

    Wk_i = Cells(Rows.Count, "A").End(xlUp).Row
    ' #comment: Detect last log line

    i = 0
    Date_i = 0
    ' #comment: Initialize counters

    Do Until i = Wk_i

        If ActiveCell.Offset(i, 0) <> "" Then
            ' #comment: Skip empty rows

            Edit_Date_Log = Left(ActiveCell.Offset(i, 0), 16)
            ' #comment: Extract full datetime string

            ActiveCell.Offset(i, 8) = Format(Edit_Date_Log, "mmdd")
            ' #comment: Store MMDD

            ActiveCell.Offset(i, 9) = Format(Edit_Date_Log, "hh:mm:ss")
            ' #comment: Store time
        End If

        ActiveCell.Offset(i, 10) = i + 1
        ' #comment: Assign row sequence number

        ActiveCell.Offset(i, 11) = 0
        ' #comment: Reset exclusion flag

        ' =============================
        ' Noise / exception filtering
        ' =============================

        If InStr(ActiveCell.Offset(i, 0), "closed") _
           Or InStr(ActiveCell.Offset(i, 0), "by uid") Then
            ActiveCell.Offset(i, 11) = 1
            ' #comment: Mark session close events
        End If

        If InStr(ActiveCell.Offset(i, 0), "pam_vas: Authentication ignored") Then
            ActiveCell.Offset(i, 11) = 1
            ' #comment: Ignore PAM authentication noise
        End If

        ' =============================
        ' Date consistency check
        ' =============================
        If Edit_Date_Prm = Format(Edit_Date_Log, "mmdd") Then
            Date_i = Date_i + 1
            ' #comment: Count valid date matches
        End If

        i = i + 1
        ' #comment: Move to next log line
    Loop

    If Date_i < 30 Then
        MsgBox "syslog生成不備の可能性あり！確認要：" & In_File_name_syslog
        ' #comment: Warn if insufficient logs
    End If

    Workbooks(In_File_name_syslog).Close SaveChanges:=False
    ' #comment: Close syslog file without saving

End Sub

Sub Sec_AIX_Main()

    Dim Loop_Neo As Integer
    ' #comment: Output row counter

    Dim Wk_Str1 As String
    ' #comment: Temporary string buffer

    Dim Wk_Before_su As String
    ' #comment: User before su

    Dim Wk_After_su As String
    ' #comment: User after su

    Dim Wk_Column As Byte
    ' #comment: Column position of substring

    Dim Tmp_Str1 As Variant
    ' #comment: Array for Split()

    Dim Wk_Pts1 As String
    ' #comment: PTS identifier

    Dim Pts_Idx As Byte
    ' #comment: Index into PTS table

    Dim i As Long
    ' #comment: Log row pointer

    Loop_Neo = 1
    ' #comment: Start writing output from row 1

    Workbooks.Open Filename:=In_Path1 & In_Path2 & In_File_name_syslog, ReadOnly:=True
    ' #comment: Open AIX syslog file

    Do Until ActiveCell = ""

        If InStr(ActiveCell, "su:") > 0 Or InStr(ActiveCell, "sudo:") > 0 Then
            ' #comment: Detect su or sudo log entries

            Edit_Date_Log = Left(ActiveCell, 16)
            ' #comment: Extract timestamp

            Edit_Date_Log_mmdd = Format(Edit_Date_Log, "mmdd")
            ' #comment: Extract MMDD

            Edit_Date_Log_hhmmss = Format(Edit_Date_Log, "hh:mm:ss")
            ' #comment: Extract time

            Wk_Column = InStr(ActiveCell, "USER=")
            ' #comment: Locate USER field

            If Wk_Column > 0 Then
                Wk_Str1 = Mid(ActiveCell, Wk_Column + 5, 12)
                ' #comment: Extract username

                Wk_Str1 = Replace(Wk_Str1, " ", "")
                ' #comment: Trim spaces

                Tmp_Str1 = Split(Wk_Str1, ":")
                ' #comment: Split at delimiter

                Wk_User = Tmp_Str1(0)
                ' #comment: Assign user
            End If

            ' ---- Write output ----
            ThisWorkbook.Sheets("NEO_Work_Sheet").Cells(Loop_Neo, 1) = "NEO_SU"
            ThisWorkbook.Sheets("NEO_Work_Sheet").Cells(Loop_Neo, 2) = Edit_Date_Log_mmdd
            ThisWorkbook.Sheets("NEO_Work_Sheet").Cells(Loop_Neo, 3) = Edit_Date_Log_hhmmss
            ThisWorkbook.Sheets("NEO_Work_Sheet").Cells(Loop_Neo, 6) = Wk_User
            ' #comment: Write parsed AIX SU event

            Call Sec_Get_Kanri_no
            ' #comment: Resolve Kanri number

            ThisWorkbook.Sheets("NEO_Work_Sheet").Cells(Loop_Neo, 8) = Kanri_No
            ThisWorkbook.Sheets("NEO_Work_Sheet").Cells(Loop_Neo, 9) = Script_File_Name
            ' #comment: Write management metadata

            Loop_Neo = Loop_Neo + 1
            ' #comment: Advance output row
        End If

        ActiveCell.Offset(1, 0).Select
        ' #comment: Move to next syslog line

    Loop

    Workbooks(In_File_name_syslog).Close SaveChanges:=False
    ' #comment: Close syslog

End Sub


Sub Write_CSV()

    Dim csvFile As String
    ' #comment: Output CSV file path

    Dim i As Long
    ' #comment: Row index

    csvFile = In_Path1 & In_Path2 & "NEOsu_" & Edit_Date_Prm & "_accounted.csv"
    ' #comment: Build output CSV filename

    Open csvFile For Output As #1
    ' #comment: Open CSV for writing

    i = 1
    Do While Sheets("NEO_Work_Csv").Cells(i, 1) <> ""

        Print #1, Sheets("NEO_Work_Csv").Cells(i, 1).Value
        ' #comment: Write one CSV line

        i = i + 1
        ' #comment: Next row
    Loop

    Close #1
    ' #comment: Close CSV file

End Sub


Sub Sec_Script_Spool()

    Dim StrFileName As String
    ' #comment: Script log filename

    Dim Tmp_Script As Variant
    ' #comment: Split filename tokens

    Erase Script_TBI_Account
    Erase Script_TBI_St_Time
    Erase Script_TBI_End_Time
    Erase Script_TBI_Kanri_No
    Erase Script_TBI_Script_File_Name
    ' #comment: Reset script metadata arrays

    Idx = 0
    ' #comment: Reset index

    StrFileName = Dir(In_Path1 & In_Path2 & "*", vbNormal)
    ' #comment: Get first script file

    Do While StrFileName <> ""

        Tmp_Script = Split(StrFileName, "_")
        ' #comment: Split filename by "_"

        If UBound(Tmp_Script) >= 4 Then
            Script_TBI_Account(Idx) = Tmp_Script(1)
            Script_TBI_St_Time(Idx) = Tmp_Script(3)
            Script_TBI_End_Time(Idx) = Format(FileDateTime(In_Path1 & In_Path2 & StrFileName), "yyyymmddhhmmss")
            Script_TBI_Kanri_No(Idx) = Tmp_Script(4)
            Script_TBI_Script_File_Name(Idx) = StrFileName
            ' #comment: Store parsed script metadata

            Idx = Idx + 1
            ' #comment: Next index
        End If

        StrFileName = Dir()
        ' #comment: Next file
    Loop

End Sub


Sub Sec_Get_Kanri_no()

    Dim ymdhms As String
    ' #comment: Log timestamp for matching

    Dim Wk_Column As Byte
    ' #comment: Column index

    Kanri_No = ""
    Script_File_Name = ""
    ' #comment: Reset outputs

    ymdhms = Format(Edit_Date_Log, "yyyymmddhhmmss")
    ' #comment: Convert log time to comparable format

    Idx = 0
    Do Until Script_TBI_Account(Idx) = ""

        If Script_TBI_Account(Idx) = Wk_User Then
            If ymdhms >= Script_TBI_St_Time(Idx) And ymdhms <= Script_TBI_End_Time(Idx) Then
                Kanri_No = Script_TBI_Kanri_No(Idx)
                Script_File_Name = Script_TBI_Script_File_Name(Idx)
                ' #comment: Matching script log found
            End If
        End If

        Idx = Idx + 1
        ' #comment: Next script entry
    Loop

End Sub

Sub SU_Count_Summary()

    Dim i As Long
    ' #comment: Loop index

    Sheets("SU_Count").Range("A1").Select
    ' #comment: Start summary output

    i = 0
    Do Until SU_Count_System(i) = ""

        ActiveCell.Offset(0, 0) = SU_Count_System(i)
        ActiveCell.Offset(1, 0) = SU_Count_Syslog(i)
        ActiveCell.Offset(2, 0) = SU_Count_X(i)
        ActiveCell.Offset(3, 0) = SU_Count_Omit(i)
        ActiveCell.Offset(4, 0) = SU_Count_Other(i)
        ' #comment: Write SU count values

        ActiveCell.Offset(0, 1).Select
        ' #comment: Move to next column

        i = i + 1
        ' #comment: Next system
    Loop

End Sub
