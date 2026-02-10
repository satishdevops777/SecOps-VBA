```vba
Sub format_changeVer() 
```
is the declaration of a VBA macro (subroutine) named format_changeVer.
```vba
Dim target As String '入力ファイルのフルパス
```
- **Meaning:** Stores the **full path** of the input file (example: `"C:\Users\Satish\Documents\data.xlsx"`).  
- **Use:** Helps VBA open or refer to a specific Excel file.  
- **Comment translation:** “Full path of input file.”

```vba
Dim target_path As String '入力ファイルのパス
```
Meaning: Stores only the folder path part of target.
Example: if target = "C:\data\report.xlsx", then target_path = "C:\data\".
Comment translation: “Path of input file.”


```vba
Dim target_name As String '入力ファイルのファイル名
```
- **Meaning:** Stores just the **filename** (without the path).  
- Example: `"report.xlsx"` or `"report.csv"`.  
- **Comment translation:** “Filename of input file.”

```vba
Dim i As Long 'index
```
Meaning: Generic loop counter or index variable (commonly used in For i = 1 To N loops).
Comment translation: “Index.”

```vba
Dim in_row As Long '入力ファイルの行カウンタ
```
- **Meaning:** Tracks the **current row number** being processed from the input file.  
- Example: increments from row 2 to last row during reading.  
- **Comment translation:** “Row counter for input file.”

 ```vba
Dim param() As String 'Paramシートのパラメタ
```
Meaning: Declares a dynamic string array used to store parameters (possibly from a “Param” worksheet).
You can later use ReDim param(n) to size it dynamically.
Comment translation: “Parameters from the Param sheet.”

```vba
Dim last_row As Long '入力ファイルの最終行
```
- **Meaning:** Stores the **last row number** in the input file that contains data.  
- Often found via `Cells(Rows.Count, 1).End(xlUp).Row`.  
- **Comment translation:** “Last row of input file.”

---

```vba
Dim wb As Workbook '出力ファイルのオブジェクト
```
Meaning: A Workbook object representing the output Excel file.
Used to reference and write results to a new or existing workbook.
Comment translation: “Object for the output file.”

```vba
Dim out_row As Long '出力ファイルの行カウンタ
```
- **Meaning:** Tracks the **current row** being written in the output file.  
- Example: starts at row 2 and increments as new data is added.  
- **Comment translation:** “Row counter for output file.”


```vba
Dim out_flg As Boolean '出力フラグ
```
Meaning: A flag variable (True/False) used to control whether or not data should be output.
Example:
If out_flg = True Then
    ' Write to output
End If
Comment translation: “Output flag.”

```vba
target = Application.GetOpenFilename("Excel ブック,*.xlsx")
```
This line is the core of the file selection process.

🧠 What It Does
Application.GetOpenFilename opens a standard Windows “Open File” dialog box.
The user can browse and select a file.
The result (the full file path) is stored in the variable target.


📁 The Filter
```vba
"Excel ブック,*.xlsx" means:
```
Only show files with the .xlsx extension.
“Excel ブック” just means “Excel Workbook” in Japanese — it’s the text shown in the dialog box.

```vba
If target = "False" Then
    MsgBox "キャンセルされました。", vbInformation
    Exit Sub
End If
```
🧠 Why This Check Exists
If the user clicks “Cancel”, the Application.GetOpenFilename function does not return a file path — it instead returns the string "False" (yes, literally the word “False”, not a Boolean value).


| Line                                                       | Function           | What Happens                           |
| ---------------------------------------------------------- | ------------------ | -------------------------------------- |
| `'入力ファイルを選択する...`                                          | Comment            | Describes purpose of next lines        |
| `target = Application.GetOpenFilename("Excel ブック,*.xlsx")` | Opens file dialog  | Lets user choose an Excel file (.xlsx) |
| `If target = "False" Then`                                 | Checks if canceled | User clicked "Cancel"                  |
| `MsgBox "キャンセルされました。"`                                     | Message box        | Shows “Canceled” message               |
| `Exit Sub`                                                 | Stops macro        | Prevents further execution             |

```vba
target_name = Dir(target)
```
target_name = "report.xlsx"

🧠 What It Does:
The built-in VBA function Dir() extracts the file name portion from a full file path.
✅ So Dir() basically removes everything before the last backslash () in the path.


target_path = Replace(target, target_name, "")

🧠 What It Does:
This line takes the full path (target) and removes the file name (target_name) from it — leaving only the directory path.

| Variable      | Value                                   |
| ------------- | --------------------------------------- |
| `target`      | `C:\Users\Satish\Documents\report.xlsx` |
| `target_name` | `report.xlsx`                           |
| `target_path` | `C:\Users\Satish\Documents\`            |

| Action                                           | Result                   |
| ------------------------------------------------ | ------------------------ |
| User selects file                                | `C:\data\sales2025.xlsx` |
| `target = "C:\data\sales2025.xlsx"`              |                          |
| `target_name = Dir(target)`                      | `"sales2025.xlsx"`       |
| `target_path = Replace(target, target_name, "")` | `"C:\data\"`             |


1) Load parameters from the Param sheet into the dynamic array param()
```vba
i = 0
in_row = 2
Do Until ThisWorkbook.Worksheets("Paramシート").Cells(in_row, 1).Value = ""
    ReDim Preserve param(i)
    param(i) = ThisWorkbook.Worksheets("Paramシート").Cells(in_row, 1).Value
    i = i + 1
    in_row = in_row + 1
Loop
```

i = 0 initializes an index for the array (param will be 0-based here).
in_row = 2 assumes parameters start on row 2 (row 1 is a header).
Do Until ... = "" loops until column A is blank (empty cell terminates input).
ReDim Preserve param(i) resizes param to hold one more element while preserving existing values.

param(i) = ...Cells(in_row,1).Value stores the cell value into param.

i = i + 1 and in_row = in_row + 1 advance counters.


2) Open the input workbook and prepare MEMBERS sheet
```vba
Workbooks.Open target
last_row = Workbooks(target_name).Worksheets("MEMBERS").Cells(Rows.Count, 2).End(xlUp).Row

Workbooks(target_name).Worksheets("MEMBERS").Range("Z1").Value = "absolute_number"
Workbooks(target_name).Worksheets("MEMBERS").Range("AA1").Value = "FLG"
```

Workbooks.Open target opens the file whose full path is in target.
last_row = ...Cells(Rows.Count,2).End(xlUp).Row finds the last used row in column B (column 2) of the MEMBERS sheet.
Practical caution: Rows.Count and Cells are not fully qualified here; if another workbook or sheet is active those references might point to the wrong sheet. Safer to fully qualify: Workbooks(target_name).Worksheets("MEMBERS").Rows.Count.
Next two lines write header labels into column Z (26) and AA (27).


3) Fill in absolute_number and FLG columns on the input sheet
```vba
in_row = 2
Do Until in_row > last_row + 1
    Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 26).Value = in_row - 1
    If Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 1).Value = "" Then
        Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 27).Value = "1"
    Else
        Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 27).Value = "2"
    End If
    in_row = in_row + 1
Loop
```
in_row = 2 starts processing data rows.

Do Until in_row > last_row + 1 loops — note this loop condition is unusual (it goes to last_row + 1). Probably intended to handle the last row but risks processing one extra row; typical pattern is Do Until in_row > last_row or For in_row = 2 To last_row.

Cells(in_row,26).Value = in_row - 1 writes a sequential number into column Z (absolute_number). For row 2 this yields 1.

The If ... Cells(in_row,1).Value = "" Then sets FLG in column AA:

If column A is blank → FLG = "1"

Else → FLG = "2"

in_row = in_row + 1 advances loop.

4) Create the output workbook and set up headers
```vba
Set wb = Workbooks.Add

wb.Worksheets(1).Name = "MEMBERS_" & Left(target_name, InStr(target_name, "00") - 1)
wb.Worksheets(1).Range("A1").Value = "Group"
... (many .Range("...").Value = "..." lines) ...
wb.Worksheets(1).Range("AC1").Value = "BU"

```
Set wb = Workbooks.Add creates a new workbook object wb to receive the output.

The sheet is renamed to "MEMBERS_" & Left(target_name, InStr(target_name, "00") - 1):

This attempts to take the portion of target_name to the left of the substring "00" — if "00" is not present, InStr returns 0, and Left(..., -1) will error. This is potentially fragile — ensure "00" exists in file names used.

The subsequent lines set many header labels in A1:T1 etc., plus Z1/AA1/AC1.

5) Loop through input details, filter by param() and write to output
```vba
out_row = 2
in_row = 2
Do Until in_row > last_row
    If Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 1).Value <> "" Then
        out_flg = False
        i = 0
        Do Until i > UBound(param)
            If Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 1).Value Like "*" & param(i) & "*" Then
                out_flg = True
            End If
            i = i + 1
        Loop

        If out_flg = True Then
            wb.Worksheets(1).Cells(out_row, 1).Value = Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 1).Value
            wb.Worksheets(1).Cells(out_row, 2).Value = Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 2).Value
            wb.Worksheets(1).Cells(out_row, 3).Value = Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 3).Value
        End If
    End If

    'detail rows handling
    If Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 1).Value = "" And _
       Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 2).Value <> "" And _
       out_flg = True Then

        If wb.Worksheets(1).Cells(out_row, 1).Value = "" Then
            wb.Worksheets(1).Cells(out_row, 1).Value = wb.Worksheets(1).Cells(out_row - 1, 1).Value
            wb.Worksheets(1).Cells(out_row, 2).Value = wb.Worksheets(1).Cells(out_row - 1, 2).Value
            wb.Worksheets(1).Cells(out_row, 3).Value = wb.Worksheets(1).Cells(out_row - 1, 3).Value
        End If

        wb.Worksheets(1).Cells(out_row, 4).Value = Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 2).Value
        wb.Worksheets(1).Cells(out_row, 5).Value = Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 3).Value
        wb.Worksheets(1).Cells(out_row, 26).Value = Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 26).Value
        wb.Worksheets(1).Cells(out_row, 27).Value = Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 27).Value
    End If

    in_row = in_row + 1
    If Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 1).Value <> "" Then
        out_row = out_row + 1
    End If
Loop
```

out_row = 2 is the write pointer into the output workbook; in_row iterates input rows.

For each row where input column A is not blank:

out_flg is set to False, then each param(i) is checked: if input col A contains any param(i) (using Like "*keyword*"), set out_flg = True.

If out_flg is True the macro copies columns 1..3 from input into output columns A..C at out_row.

Then there is another conditional block for “detail rows”:

If input col A is blank and input col B is not blank and out_flg = True (i.e., this is a continuation/detail row belonging to the previously matched group), the macro:

If output A (Group) is blank, it copies the previous row’s A,B,C into current output A,B,C (so detail rows inherit group header).

Writes input col B to output col D, input col C to output col E, and copies absolute_number and FLG (cols 26/27) to output cols 26/27.

After processing, in_row is incremented. Then the macro checks the next input row: if its col A is not blank, out_row is incremented. That logic controls when to move to a new output row.

6) Format the output workbook
```vba
wb.Worksheets(1).Range("A1:F1").Interior.Color = 16777062 'light blue
wb.Worksheets(1).Range("G1:I1").Interior.Color = 65535   'yellow
wb.Worksheets(1).Range("J1:T1").Interior.Color = 5296274  'green
wb.Worksheets(1).Range("A1:T" & out_row - 1).Font.Bold = True
wb.Worksheets(1).Range("A1:T" & out_row - 1).Borders.LineStyle = xlContinuous
wb.Worksheets(1).Columns("A:AC").AutoFit
wb.Worksheets(1).Columns("A:AC").AutoFilter
```
Sets header background colors for groups of columns (color codes used directly).

Makes header font bold up to out_row - 1 (note: if out_row ended up as 2 and no rows written, this might produce A1:T1 which is OK).

Adds borders to the range.

AutoFits columns A:AC and enables AutoFilter on them.


7) Add dropdown lists (data validation) to columns G and H

```vba
With wb.Worksheets(1).Range("G2:G" & out_row)
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="Retain"
    .IgnoreBlank = True
    .InCellDropdown = True
    .ShowInput = True
    .ShowError = True
End With

With wb.Worksheets(1).Range("H2:H" & out_row)
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="Remove"
    .IgnoreBlank = True
    .InCellDropdown = True
    .ShowInput = True
    .ShowError = True
End With
```
This block attempts to apply a validation list to columns G and H for rows 2..out_row.

Formula1:="Retain" sets a single-item list containing "Retain". Same for "Remove" in column H.

If the intention is to let users choose between "Retain" and "Remove" in the same column, Formula1 should be "Retain,Remove" or a reference to a range with both items.

Implementation note: .Add is a method of the Validation object — calling .Add directly on the Range can work because VBA maps it, but some prefer Range.Validation.Add ... and then set .Validation properties.

8) Remove extra sheets in the new workbook
```vba
Do While wb.Worksheets.Count > 1
    wb.Worksheets(2).Delete
Loop
```

Workbooks.Add usually creates multiple sheets (depending on Excel settings). This loop removes all sheets except the first by always deleting sheet index 2 until only one remains.

Note: Deleting sheets can prompt the user unless Application.DisplayAlerts = False is set prior to deletion and then restored afterwards.
Option Explicit

9) Save and close output, save and close input
```
wb.SaveAs target_path & Replace(target_name, ".xlsx", "") & "_Member.xlsx"

wb.Close

Workbooks(target_name).Save
Workbooks(target_name).Close

MsgBox "処理が完了しました。", vbInformation

````
Saves the wb workbook into the same folder as the input (target_path), with name: original filename without .xlsx plus "_Member.xlsx".

Example: if target_name = "data00.xlsx", saved file becomes data00_Member.xlsx.

Closes the output workbook.

Saves and closes the input workbook.

Shows a completion message box ("Processing is complete.").

```vba
Sub format_changeVer()

    Dim target As String '入力ファイルのフルパス
    Dim target_path As String '入力ファイルのパス
    Dim target_name As String '入力ファイルのファイル名
    Dim i As Long 'index
    Dim in_row As Long '入力ファイルの行カウンタ
    Dim param() As String 'Paramシートのパラメタ
    Dim last_row As Long '入力ファイルの最終行
    Dim wb As Workbook '出力ファイルのオブジェクト
    Dim out_row As Long '出力ファイルの行カウンタ
    Dim out_flg As Boolean '出力フラグ

    '入力ファイルを選択する(未選択はキャンセル扱い)
    target = Application.GetOpenFilename("Excel ブック,*.xlsx")
    If target = "False" Then
        MsgBox "キャンセルされました。", vbInformation
        Exit Sub
    End If

    '入力ファイルのパス、入力ファイルのファイル名を設定する
    target_name = Dir(target)
    target_path = Replace(target, target_name, "")

    'Paramシートのパラメタを動的配列へ格納する
    i = 0
    in_row = 2
    Do Until ThisWorkbook.Worksheets("Paramシート").Cells(in_row, 1).Value = ""
        ReDim Preserve param(i)
        param(i) = ThisWorkbook.Worksheets("Paramシート").Cells(in_row, 1).Value
        i = i + 1
        in_row = in_row + 1
    Loop

    '入力ファイルを開く
    Workbooks.Open target

    '入力ファイルの最終行を設定する
    last_row = Workbooks(target_name).Worksheets("MEMBERS").Cells(Rows.Count, 2).End(xlUp).Row

    '入力ファイルのヘッダを設定する
    Workbooks(target_name).Worksheets("MEMBERS").Range("Z1").Value = "absolute_number"
    Workbooks(target_name).Worksheets("MEMBERS").Range("AA1").Value = "FLG"

    '入力ファイルの明細を設定する
    in_row = 2
    Do Until in_row > last_row + 1
        Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 26).Value = in_row - 1
        If Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 1).Value = "" Then
            Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 27).Value = "1"
        Else
            Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 27).Value = "2"
        End If
        in_row = in_row + 1
    Loop

    '出力ファイルを作成する
    Set wb = Workbooks.Add

    '出力ファイルのヘッダを設定する
  
    wb.Worksheets(1).Name = "MEMBERS_" & Left(target_name, InStr(target_name, "00") - 1)
    wb.Worksheets(1).Range("A1").Value = "Group"
    wb.Worksheets(1).Range("B1").Value = "Group Owners"
    wb.Worksheets(1).Range("C1").Value = "Group Recert"
    wb.Worksheets(1).Range("D1").Value = "Account"
    wb.Worksheets(1).Range("E1").Value = "Account Name"
    wb.Worksheets(1).Range("F1").Value = "Server Name"
    wb.Worksheets(1).Range("G1").Value = "Retain Access"
    wb.Worksheets(1).Range("H1").Value = "Remove Access"
    wb.Worksheets(1).Range("I1").Value = "Recertifier"
    wb.Worksheets(1).Range("J1").Value = "Application Name"
    wb.Worksheets(1).Range("K1").Value = "Service Account"
    wb.Worksheets(1).Range("L1").Value = "プライマリーオーナー"
    wb.Worksheets(1).Range("M1").Value = "再認定対象者"
    wb.Worksheets(1).Range("N1").Value = "XID/ID"
    wb.Worksheets(1).Range("O1").Value = "Function/所属"
    wb.Worksheets(1).Range("P1").Value = "Product/Tower"
    wb.Worksheets(1).Range("Q1").Value = "Sub Product"
    wb.Worksheets(1).Range("R1").Value = "Recertifier"
    wb.Worksheets(1).Range("S1").Value = "RSM(ID)"
    wb.Worksheets(1).Range("T1").Value = "RSM(Name)"
    wb.Worksheets(1).Range("Z1").Value = "absolute_number"
    wb.Worksheets(1).Range("AA1").Value = "FLG"
    wb.Worksheets(1).Range("AC1").Value = "BU"

    '出力ファイルの明細を設定する
    out_row = 2
    in_row = 2
    Do Until in_row > last_row
        If Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 1).Value <> "" Then
            out_flg = False
            'Paramシートのパラメタ数分以下の処理を繰り返す
            i = 0
            Do Until i > UBound(param)
                If Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 1).Value Like "*" & param(i) & "*" Then
                    out_flg = True
                End If
                i = i + 1
            Loop

            If out_flg = True Then
                wb.Worksheets(1).Cells(out_row, 1).Value = Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 1).Value
                wb.Worksheets(1).Cells(out_row, 2).Value = Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 2).Value
                wb.Worksheets(1).Cells(out_row, 3).Value = Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 3).Value
            End If
        End If

        '明細処理
        If Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 1).Value = "" And _
           Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 2).Value <> "" And _
           out_flg = True Then

            If wb.Worksheets(1).Cells(out_row, 1).Value = "" Then
                wb.Worksheets(1).Cells(out_row, 1).Value = wb.Worksheets(1).Cells(out_row - 1, 1).Value
                wb.Worksheets(1).Cells(out_row, 2).Value = wb.Worksheets(1).Cells(out_row - 1, 2).Value
                wb.Worksheets(1).Cells(out_row, 3).Value = wb.Worksheets(1).Cells(out_row - 1, 3).Value
            End If

            wb.Worksheets(1).Cells(out_row, 4).Value = Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 2).Value
            wb.Worksheets(1).Cells(out_row, 5).Value = Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 3).Value
            wb.Worksheets(1).Cells(out_row, 26).Value = Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 26).Value
            wb.Worksheets(1).Cells(out_row, 27).Value = Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 27).Value
        End If

        in_row = in_row + 1
        If Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 1).Value <> "" Then
            out_row = out_row + 1
        End If
    Loop

    '出力ファイルのフォーマット編集する
    wb.Worksheets(1).Range("A1:F1").Interior.Color = 16777062 '水色
    wb.Worksheets(1).Range("G1:I1").Interior.Color = 65535 '黄色
    wb.Worksheets(1).Range("J1:T1").Interior.Color = 5296274 '緑色
    wb.Worksheets(1).Range("A1:T" & out_row - 1).Font.Bold = True '太字
    wb.Worksheets(1).Range("A1:T" & out_row - 1).Borders.LineStyle = xlContinuous '格子
    wb.Worksheets(1).Columns("A:AC").AutoFit '自動調整
    wb.Worksheets(1).Columns("A:AC").AutoFilter 'フィルターをかける

    'ドロップダウンリストの追加
    With wb.Worksheets(1).Range("G2:G" & out_row)
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="Retain"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With

    With wb.Worksheets(1).Range("H2:H" & out_row)
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="Remove"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With

    '不要なシートを削除
    Do While wb.Worksheets.Count > 1
        wb.Worksheets(2).Delete
    Loop

    '出力ファイルを保存する
    wb.SaveAs target_path & Replace(target_name, ".xlsx", "") & "_Member.xlsx"

    '出力ファイルを閉じる
    wb.Close

    '入力ファイルを保存する
    Workbooks(target_name).Save

    '入力ファイルを閉じる
    Workbooks(target_name).Close

    MsgBox "処理が完了しました。", vbInformation

End Sub

```
