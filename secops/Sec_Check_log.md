## 1️⃣ What is the USE of this VBA program?
 ### This macro is a security log analysis tool.
-  It automatically checks Linux server login logs (last and accounted CSV files)
-  It detects suspicious or meaningful logins
-  It counts & classifies accounts
-  It produces a formatted Excel report

## 2️⃣ Prerequisites (VERY IMPORTANT)
- Before running this macro, the following must exist:

### ✅ Excel Sheets (Mandatory)
### 1. Param_Sheet_last_su
    - This is the control sheet (input parameters).
    - Each row represents one server / one log set.

Columns used:

| Column | Meaning                 |
| ------ | ----------------------- |
| A      | Enable flag (Y / y / ✓) |
| B      | Business code           |
| C      | Check pattern           |
| D      | Input folder path       |
| E      | Sub folder              |
| F      | LAST log filename       |
| G      | ACCOUNTED log filename  |
| H      | OS type                 |
| I      | OS version              |

### 2. Result_last
    - This is the output report sheet
    - The macro clears and rewrites this sheet every run.

### ✅ Input Files (External CSVs)

- For each enabled row, these files must exist:
  - last log CSV
  - accounted log CSV
 
Path used:
```
In_Path1 + In_Path2 + FileName

Example:
C:\Logs\Server01\last.csv
C:\Logs\Server01\accounted.csv
```

### ✅ Excel Settings
- Macros enabled
- VBA access allowed
- Files encoded as UTF-8 (65001)

```vb
Option Explicit                 ' Forces all variables to be declared (prevents bugs)

Public result(199) As String    ' Array to store up to 200 LAST log log lines
Public result_cnt As Integer   ' Counter to track number of LAST log entries stored

Dim loop_i As Integer           ' Loop counter for Param_Sheet rows
Dim Code As String              ' Business / server code
Dim Code_Color As Long          ' Background color of the business code cell
Dim Chk_Ctl As String           ' Execution control flag (Y / y / ✓)
Dim Check_Ptn As String         ' Pattern used for log checking
Dim In_Path1 As String          ' Base folder path for log files
Dim In_Path2 As String          ' Sub folder path for log files
Dim In_File_name_last As String ' LAST log file name
Dim In_File_name_Accounted As String ' ACCOUNTED log file name
Dim Chk_OS As String            ' Operating system name
Dim Chk_OS_Ver As String        ' Operating system version
Dim Wk_Activecell As String     ' Value used to detect end of parameter rows
Dim Wk_Str As String            ' Temporary string to hold one log line

Worksheets("Result_last").Cells.Clear   ' Clear previous execution results

Wk_Activecell = "9999"          ' Dummy value to enter loop
loop_i = 2                      ' Start reading Param_Sheet from row 2
Application.ScreenUpdating = False ' Disable screen refresh to improve speed

Do Until Wk_Activecell = ""     ' Continue until empty row is found

Chk_Ctl = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 2)
                                ' Read execution control flag from column B
If Chk_Ctl = "Y" Or Chk_Ctl = "y" Or Chk_Ctl = "✓" Then
                                ' Process only enabled rows


Code = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 1)
                                ' Read business/server code

Code_Color = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 1).Interior.Color
                                ' Capture color of business code cell

Check_Ptn = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 3)
                                ' Read check pattern

In_Path1 = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 4)
                                ' Read base path

In_Path2 = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 5)
                                ' Read sub path

In_File_name_last = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 6)
                                ' Read LAST log filename

In_File_name_Accounted = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 7)
                                ' Read ACCOUNTED log filename

Chk_OS = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 8)
                                ' Read OS name

Chk_OS_Ver = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 9)
                                ' Read OS version


Workbooks.OpenText _
    Filename:=In_Path1 & In_Path2 & In_File_name_last, _
    Origin:=65001, _
    DataType:=xlDelimited
                                ' Open LAST log CSV using UTF-8 encoding

End_row_last = Cells(1, 1).SpecialCells(xlLastCell).Row
                                ' Get last used row in LAST log file

ReDim result(End_row_last - 1)  ' Resize result array to match log size
result_cnt = 0                 ' Initialize result counter


cnt_i = 1                       ' Start reading from first row

Do Until cnt_i > End_row_last   ' Loop through all rows in LAST log


Wk_Str = Cells(cnt_i, 1)        ' Read one line from LAST log


If cnt_i = 1 Then               ' Check first line only
    If Left(Wk_Str, 3) <> "[*-" Then
        MsgBox "Last log format may be invalid: " & In_File_name_last
                                ' Show warning if log format looks wrong
    End If
End If

result(result_cnt) = Wk_Str     ' Store log line into result array
result_cnt = result_cnt + 1    ' Increment result counter
cnt_i = cnt_i + 1              ' Move to next row


Workbooks(In_File_name_last).Close
                                ' Close LAST log workbook

Workbooks.Open _
    Filename:=In_Path1 & In_Path2 & In_File_name_Accounted, _
    ReadOnly:=True
                                ' Open ACCOUNTED log file in read-only mode

End_row_Accounted = 0           ' Initialize accounted row count

If Cells(1, 1) <> "" Then
    End_row_Accounted = Cells(1, 1).SpecialCells(xlLastCell).Row
                                ' Get last row of ACCOUNTED log
End If


judge_Accounted = ""            ' Reset judgement flag

Do Until ActiveCell = ""        ' Loop through ACCOUNTED entries

If ActiveCell = "root" _
   Or ActiveCell = "tdisop" _
   Or ActiveCell = "qscanner" _
   Or ActiveCell = "wasroot" Then
                                ' Ignore known system accounts
Else
    judge_Accounted = "有"      ' Mark suspicious account found
End If

ActiveCell.Offset(1, 0).Select ' Move to next account
Loop

Call sec_Count_paste( _
    Code, _
    Code_Color, _
    In_Path2, _
    In_File_name_last, _
    In_File_name_Accounted, _
    End_row_Accounted, _
    judge_Accounted)
                                ' Write collected data into Result_last sheet

loop_i = loop_i + 1             ' Move to next parameter row
Wk_Activecell = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 1)
                                ' Check if next row exists

Loop                            ' Continue parameter sheet loop

Call Sec_Check_Exciusion        ' Perform final exclusion and summary logic

End Sub                         ' End of Sec_Check_log procedure
```
