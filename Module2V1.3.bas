VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub CommandButton1_Click()


'Open and sort DBF
Dim DBFSelect As Variant
Dim fd As FileDialog
Dim wsPline As Worksheet
'Rows Copy
Dim mstRows As Long
Dim wsInstDate As Worksheet
Dim wbCompRes As Workbook, wbPline As Workbook
Dim i As Integer
Dim rowCounter As Integer
'Copy and clear master wkst
Dim InstDate As Variant
Dim yr As Variant
Dim Wkb1 As Workbook

'Activate the File Dialog Open------------------------------------------------------------

Set fd = Application.FileDialog(msoFileDialogOpen)
    With fd
        .Title = "Please select the pline/transect intersect DBF file."
        .Filters.Clear
        .AllowMultiSelect = False
        .Filters.Add "Database File ", "*.dbf"
        If .Show = -1 Then
'Save the File name as DBFSelect variable and Opens
        DBFSelect = .SelectedItems(1)
        Set wbPline = Workbooks.Open(DBFSelect)
        Workbooks.Open (DBFSelect)
        Else
            Exit Sub
        End If
    End With

'Sort Pline DBF by ID_1

wbPline.Worksheets(1).Range("A1:E701").Select
With wbPline.Worksheets(1).Sort
    .SortFields.Clear
    .SortFields.Add Key:=Range("B2:B701"), SortOn:=xlSortOnValues, Order:=xlAscending, _
    DataOption:=xlSortNormal
    .SetRange Range("A1:E701")
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
'End of Open and Sort


' Copy Master Wkst and clear Master worksheet---------------------------------------------------------------
'Prompt user for Instance Date and save as variable Wkb1
InstDate = InputBox("Please type the instance date for this analysis (YYYYMMDD).")
If InstDate = "" Then
    wbPline.Close SaveChanges:=False
    Exit Sub
Else
    yr = Left(InstDate, 4)
    Set Wkb1 = Workbooks("Computation_result" & CStr(yr) & ".xls")
End If
'Copy Master Workbook and Rename and move to front
    With Wkb1.Worksheets("Master Wkst")
        .Copy Before:=Wkb1.Worksheets(1)
    End With
Sheets("Master Wkst (2)").Select
Sheets("Master Wkst (2)").Name = InstDate
Set wsInstDate = Wkb1.Worksheets(InstDate)
   
    
'Transfer DBF File coords to Master Worksheet ---------------------------------------------------
'Set wsMstr = ThisWorkbook.Worksheets("Master Wkst")
'wsMstr.Activate
mstRows = wsInstDate.Cells(Rows.Count, 1).End(xlUp).Row
'If statment to check if column a1 in Master Wkst = Column b2 in pline file
'If true, copies rows into Master Wkst, If false, skips row
rowCounter = 2
MsgBox ("Please Wait while files are copied.")
Application.ScreenUpdating = False
For i = 4 To mstRows
    If wsInstDate.Cells(i, 1).Value <> wbPline.Worksheets(1).Cells(rowCounter, 2).Value Then
    Else
    
    
    
'Out of Range! I am not sure if its an object issue with the Workbook vs Worksheet or not.
' Everything above this works fine.
    
    
        wbPline.Worksheet(1).Range(Cells(rowCounter, 4), Cells(rowCounter, 5)).Copy _
        Destination:=Workbooks("Computation_results2017.xls").Worksheet("Master Wkst").Range(Cells(i, 6), Cells(i, 7))
        'ActiveSheet.Range(Cells(rowCounter, 4), Cells(rowCounter, 5)).Copy
        'wsMstr.Activate
        'wsMstr.Cells(i, 6).Select
        'ActiveSheet.Paste
        rowCounter = rowCounter + 1
    End If
Next
Application.ScreenUpdating = True
wbPline.Close


End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
