Attribute VB_Name = "Main"
    'Following sheet variables are global and initialised in Open Workbook event
    Public WB As Workbook
    'Consolidated Support Stats
    Public WS_CSS As Worksheet
    'Main Data
    Public WS_DA As Worksheet
    'Raw Data
    Public WS_RD As Worksheet
    'Home
    Public WS_HM As Worksheet
    'Consolidated Performance Audit
    Public WS_CPA As Worksheet
    'Date of quarters will be stored here, accessible globally
    Public quarters(14, 1) As Variant

    Public StartTime As Double
    Public SecondsElapsed As Double
Sub pOpenApp()
'========================================================================================================
' pOpenApp
' -------------------------------------------------------------------------------------------------------
' Purpose of this programm : Execution starts with this procedure
'
' Author : Subhankar Paul, 9th January, 2018
' Notes  : N/A
' Parameters : N/A
' Returns : N/A
' ---------------------------------------------------------------
' Revision History
'
'========================================================================================================

    Set WB = ActiveWorkbook
    Set WS_CSS = WB.Sheets("Consolidated Support Stats")
    Set WS_DA = WB.Sheets("MainData")
    Set WS_RD = WB.Sheets("Raw Data")
    Set WS_HM = WB.Sheets("Home")
    Set WS_CPA = WB.Sheets("Consolidated Performance Audit")

On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Call InputDate

'    Call pCleanDB
    
    Call mainDataStaging
    
    Call CreateUniqueList
    
ErrorHandler:
 
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    ' Determine how many seconds this code will take to run
    SecondsElapsed = Round(Timer - StartTime, 2)
    'Notify user in seconds
    MsgBox "Application ran successfully in " & SecondsElapsed & " seconds", vbInformation
End Sub

Sub InputDate()
'========================================================================================================
' InputDate
' ------------------------------------------------------------------------------------------------------
' Purpose of this Function : Taking date input from home page as quarter
'
' Author : Subhankar Paul | 9th January, 2018
' Notes  : Procedure can take dates even from blank quarters in between
' Parameters :
' Returns :
' ---------------------------------------------------------------
' Revision History
'
'========================================================================================================
    Dim c As Integer
    Dim i As Integer
    
    WS_HM.Select
    c = 0
    For i = 5 To 33
        If (Cells(i, 4).Value <> "" And Cells(i, 6).Value <> "") Then
            quarters(c, 0) = Cells(i, 4).Value
            quarters(c, 1) = Cells(i, 6).Value
            c = c + 1
        End If
        i = i + 1
    Next i
End Sub

