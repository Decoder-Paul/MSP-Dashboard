Attribute VB_Name = "Staging"
Sub mainDataStaging()

Dim WB As Workbook
Dim WS_CR As Worksheet
Dim WS_DA As Worksheet
Dim sCr As String
Dim sDa As String
Dim CRlro As Long
Dim DAlro As Long

sCr = "Consolidated Report"
sDa = "MainData"

Set WB = ActiveWorkbook
Set WS_CR = WB.Sheets(sCr)
Set WS_DA = WB.Sheets(sDa)

WB.Sheets(sDa).Activate
Sheets(sDa).Select

DAlro = WS_DA.Cells(WS_DA.Rows.Count, "A").End(xlUp).Row
If DAlro >= 1 Then
    WS_DA.Range("A1:Z" & (DAlro)).Clear
End If

WB.Sheets(sCr).Activate
Sheets(sCr).Range("A1").Select

CRlro = WS_CR.Cells(WS_CR.Rows.Count, "A").End(xlUp).Row

WS_CR.Range(Cells(R + 1, C + 2), Cells(R + CRlro, C + 22)).Copy
Sheets(sDa).Select
WS_DA.Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False

WS_DA.Range(Cells(R + 1, C + 1), Cells(R + 1, C + 21)).Interior.Color = RGB(46, 139, 87)
WS_DA.Range(Cells(R + 1, C + 1), Cells(R + 1, C + 22)).RowHeight = 30
WS_DA.Range(Cells(R + 1, C + 1), Cells(R + 1, C + 22)).VerticalAlignment = xlCenter

Range(Cells(R + 1, C + 1), Cells(R + DAlro, C + 22)).Select
With Selection
    .Columns.AutoFit
End With

For Each BI In Array(xlEdgeTop, xlEdgeLeft, xlEdgeBottom, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)
    With Range(Cells(R + 1, C + 1), Cells(R + DAlro, C + 22)).Borders(BI)
         .Weight = xlThin
         .Color = RGB(148, 138, 84)
    End With
 Next BI
 
 
 
End Sub
