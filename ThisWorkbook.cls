VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Set WB = ActiveWorkbook
    Set WS_CSS = WB.Sheets("Consolidated Support Stats")
    Set WS_DA = WB.Sheets("MainData")
    Set WS_RD = WB.Sheets("Raw Data")
    Set WS_HM = WB.Sheets("Home")
    Set WS_CPA = WB.Sheets("Consolidated Performance Audit")
End Sub
