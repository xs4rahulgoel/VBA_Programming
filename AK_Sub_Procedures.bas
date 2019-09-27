Attribute VB_Name = "AK_Sub_Procedures"
Sub AK_Products_Price()
'
' Akamai_Products Macro

    ActiveCell.Select
    ActiveCell.FormulaR1C1 = "Product "
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "ION"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "CONA"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "AQUA"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "IP"
    ActiveCell.Offset(-4, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Price"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "5500"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "7000"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "20000"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "24000"
    ActiveCell.Offset(1, 0).Range("A1").Select
End Sub

Sub AK_Products_Price2()

'PROGRAM TO DISPLAY PRODUCTS AND PRICE


ActiveCell.Offset(0, 0) = "Products1"
ActiveCell.Offset(1, 0) = "Products2"
ActiveCell.Offset(2, 0) = "Products3"
ActiveCell.Offset(3, 0) = "Products4"
ActiveCell.Offset(4, 0) = "Products5"
ActiveCell.Offset(5, 0) = "Products6"

With ActiveCell
.Offset(0, 1) = 100
.Offset(1, 1) = 200
.Offset(2, 1) = 300
.Offset(3, 1) = 400
.Offset(4, 1) = 500
.Offset(5, 1) = 600
End With

End Sub

Sub AK_Format_Data()
'prg to format selected data

With Selection.Font
.Bold = True
.Name = "arial"
.Size = 14
.Color = vbRed
End With

' Adjustcloandrows

End Sub

Sub AK_Adjust_columns_rows()
'prg to manage the width of column and rows
Cells.Select
Selection.Columns.AutoFit
Selection.Rows.AutoFit
ActiveSheet.Range("a1").Select

End Sub
Sub AK_Chart()
'Prg to generate the chart of selected data

Charts.Add
ActiveChart.ChartType = xl3DColumnClustered
ActiveChart.ApplyDataLabels xlDataLabelsShowValue
ActiveChart.HasDataTable = True

Dim ChTitle As String, XTitle As String, yTitle As String

ChTitle = InputBox("Enter Chart Title", "Chart Title", "ABC")
XTitle = InputBox("Enter X-Axis Title", "X-Axis", "XYZ")
yTitle = InputBox("Enter Y-Axis Title", "Y-Axis", "XYZ")

'Specifying the chart title
With ActiveChart
.HasTitle = True
.ChartTitle.Text = ChTitle
End With

'Specifying the x-axis title
With ActiveChart.Axes(xlCategory)
.HasTitle = True
.AxisTitle.Text = XTitle
End With

'Specifying the x-axis title
With ActiveChart.Axes(xlValue)
.HasTitle = True
.AxisTitle.Text = yTitle
End With

End Sub


Sub AK_Pie_Chart()
'Prg to generate the chart of selected data

Dim shtname As String

shtname = ActiveSheet.Name

Charts.Add
ActiveChart.ChartType = xl3DPie
ActiveChart.ApplyDataLabels xlDataLabelsShowValue

Dim ChTitle As String

ChTitle = InputBox("Enter Chart Title", "Chart Title", "ABC")

'Specifying the chart title
With ActiveChart
.HasTitle = True
.ChartTitle.Text = ChTitle
End With

ActiveChart.Location xlLocationAsObject, shtname

End Sub

Sub AK_Hide_Sheets()
'Prg to hide the worksheets

Dim ws As Worksheet
Dim shtname As String
shtname = ActiveSheet.Name
 
On Error Resume Next

For Each ws In ActiveWorkbook.Worksheets
    If ws.Name <> shtname Then
        ws.Visible = xlSheetVeryHidden
    End If
Next ws

Err.Clear 'Stores and clears the error information
On Error GoTo 0  'Disables all error handlers

End Sub


Sub AK_UnHide_Sheets()
Attribute AK_UnHide_Sheets.VB_ProcData.VB_Invoke_Func = " \n14"
'Prg to unhide the worksheets

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
    ws.Visible = xlSheetVisible
Next ws

End Sub

Sub AK_GetFileName()

'Prg to read the files name and display the same on the work sheet
Dim myFolder As String, myFile As String
myFolder = "/Users/rkumar/Documents/ExcelTraining/DataFromTraining/TrainingFile/Data"

myFile = Dir(myFolder & "/*.*")

Dim count As Integer
count = 0

Do While myFile <> ""
ActiveCell.Offset(count, 0) = myFile
count = count + 1
myFile = Dir
Loop

End Sub

Sub AK_ConsData_From_Diff_Files_inOneFileOneSheet()

'Prg to consolidate data from different file in one file one sheet.

Dim myFolder As String, myFile As String
myFolder = "/Users/rkumar/Documents/ExcelTraining/DataFromTraining/TrainingFile/Data"
myFile = Dir(myFolder & "/*.*")

Workbooks.Add
ActiveWorkbook.SaveAs "/Users/rkumar/Desktop/Consolidated.xlsx"

Dim count As Integer, SrcFile As String, x As Integer, y As Integer

count = 0

Do While myFile <> ""

SrcFile = myFolder & "/" & myFile 'getting the complete file path
Workbooks.Open SrcFile 'Opening the source file
wbname = ActiveWorkbook.Name 'Getting the current filename.
x = ActiveSheet.UsedRange.Rows.count
y = ActiveSheet.UsedRange.Columns.count

If count = 0 Then
ActiveSheet.Range("a1", Cells(x, y)).Select
Else
ActiveSheet.Range("a2", Cells(x, y)).Select
End If
Selection.Copy

Windows("Consolidated.xlsx").Activate
x = ActiveSheet.UsedRange.Rows.count
If count <> 0 Then
x = x + 1
End If

ActiveSheet.Range("a" & x).Select
ActiveSheet.Paste
ActiveWorkbook.Save
Windows(wbname).Close

myFile = Dir 'set the focus back to folder
count = count + 1
Loop

End Sub

Sub AK_ConsData_From_Diff_Files_inOneFileMultipleSheet()

'Prg to consolidate data from different file in one file one sheet.

Dim myFolder As String, myFile As String
myFolder = "/Users/rkumar/Documents/ExcelTraining/DataFromTraining/TrainingFile/Data"
myFile = Dir(myFolder & "/*.*")

Workbooks.Add
ActiveWorkbook.SaveAs "/Users/rkumar/Desktop/Consolidated.xlsx"

Dim count As Integer, SrcFile As String, x As Integer, y As Integer

count = 0

Do While myFile <> ""

SrcFile = myFolder & "/" & myFile 'getting the complete file path
Workbooks.Open SrcFile 'Opening the source file
wbname = ActiveWorkbook.Name 'Getting the current filename.
x = ActiveSheet.UsedRange.Rows.count
y = ActiveSheet.UsedRange.Columns.count

ActiveSheet.Range("a1", Cells(x, y)).Select
Selection.Copy

Windows("Consolidated.xlsx").Activate
If count <> 0 Then
ActiveWorkbook.Sheets.Add
End If

ActiveSheet.Name = wbname
ActiveSheet.Paste
ActiveWorkbook.Save

Windows(wbname).Close

myFile = Dir 'set the focus back to folder
count = count + 1
Loop

End Sub

Sub AK_SortSheets()
Dim i As Integer, j As Integer

For i = 1 To Sheets.count
    For j = 1 To Sheets.count - 1
        If StrComp(Sheets(j).Name, Sheets(j + 1).Name) < 0 Then
            Sheets(j).Move after:=Sheets(j + 1)
        End If
    Next j
Next i

End Sub

Sub AK_FilterData()
    
Dim shtname As String
shtname = ActiveSheet.Name

    'Prg to filter data,copythe filtered data and paste in the new sheet
Range("b:b").AdvancedFilter Action:=xlFilterCopy, copytorange:=Range("g1"), Unique:=True

ActiveSheet.Range("g:g").Select
Selection.Cut
Sheets.Add
ActiveSheet.Paste
ActiveSheet.Name = "Unique"

Dim x As Integer, count As Integer, cellvalue As String

x = ActiveSheet.UsedRange.Rows.count
MsgBox x


For count = 2 To x
    cellvalue = ActiveSheet.Range("a" & count).Value
    Sheets(shtname).Select
    ActiveSheet.Rows("1:1").AutoFilter field:=2, Criteria1:=cellvalue
    ActiveSheet.AutoFilter.Range.Copy
    Sheets.Add
    ActiveSheet.Paste
    ActiveSheet.Range("a1").Select
    ActiveSheet.Name = cellvalue
    Sheets(shtname).Select
    ActiveSheet.ShowAllData
    Sheets("Unique").Select
Next count

Application.DisplayAlerts = False
ActiveSheet.Delete
Application.DisplayAlerts = True

Sheets(shtname).Select
ActiveSheet.AutoFilterMode = False
ActiveSheet.Range("a1").Select

End Sub


Sub AK_Cons_Filter_Sort()

AK_ConsData_From_Diff_Files_inOneFileOneSheet
AK_FilterData
AK_SortSheets

End Sub

Sub AK_PivotTable()
'Prg to generate the pivot table of selected data

Dim x As Integer, y As Integer

x = ActiveSheet.UsedRange.Rows.count
y = ActiveSheet.UsedRange.Columns.count

ActiveSheet.Range("a1", Cells(x, y)).Select

Dim RangeName As String

RangeName = "MyData"
Selection.Name = RangeName

Dim Pc As PivotCache
Set Pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, RangeName)
Sheets.Add
Dim Pt As PivotTable
Set Pt = Pc.CreatePivotTable(Range("a1"), "Pt1")

Dim Pf1 As PivotField, Pf2 As PivotField
Set Pf1 = Pt.PivotFields("Department")
Set Pf2 = Pt.PivotFields("Employee Name")
Pf1.Orientation = xlRowField
Pf2.Orientation = xlDataField
AK_Pie_Chart

End Sub


Sub AK_MSWordInteraction()

Selection.Copy
Dim AppWord As Word.Application
Set AppWord = CreateObject("word.Application")
AppWord.Visible = True

AppWorld.Document.Add
AppWorld.Selection.Paste

Application.CutCopyMode = False

End Sub

Sub AK_SendMail()
'ActiveWorkbook.SendMail Recipients:="rahul.k.goel@gmail.com", Subject:="Test12345"
ActiveWorkbook.SendMail Recipients:="", Subject:="Test12345"

End Sub

Sub AK_Email_Address()



End Sub

