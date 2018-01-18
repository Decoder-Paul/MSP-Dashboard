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
    'This is the no. of quarters
    Public c As Integer
    
    'Date of report
    Public DateOfreport As Date

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
    StartTime = Timer
    Set WB = ActiveWorkbook
    Set WS_CSS = WB.Sheets("Consolidated Support Stats")
    Set WS_DA = WB.Sheets("MainData")
    Set WS_RD = WB.Sheets("Raw Data")
    Set WS_HM = WB.Sheets("Home")
    Set WS_CPA = WB.Sheets("Consolidated Performance Audit")
    
    'Date of report taking from Home sheet in L column
     DateOfreport = WS_HM.Cells(5, 12).Value

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Call InputDate

    Call pCleanDB
    Call QtrReplication
    
    Call mainDataStaging
    
    Call CreateUniqueList
    
    Call teamsDashboard
    
     
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

Sub teamsDashboard()
'========================================================================================================
' Main Data for Staging
' -------------------------------------------------------------------------------------------------------
' Purpose   :   It's responsible to call other procedure for generating report team wise
'
' Author    :   Shambhavi B M, 10th January, 2018
' Notes     :   N/A
'
' Parameter :   N/A
' Returns   :   N/A
' -------------------------------------------------------------------------------------------------------
' Revision History
'
'========================================================================================================

    Dim DAlro As Long
    Dim a As Long
    Dim team As String
    
    WS_DA.Activate

    'getting no. teams from main data
    DAlro = WS_DA.Cells(WS_DA.Rows.Count, "V").End(xlUp).Row
    
    'Calling ReplicateMainSheet method for each team
    For i = 2 To DAlro
        WS_DA.Activate
        'getting team name from main data
        team = Cells(i, 22).Value
        Call agingCount(team)
        'generating the dashboard for each team and quarterwise
        For j = 0 To c - 1
            Call ticketCount(team, j)
        Next j
        'replicating the team's dashboard
        Call ReplicateMainSheet(team)
        'cleansing of CSS sheet after replication
        'Active Ticket's stat table
        WS_CSS.Range("D5:R9").ClearContents
        WS_CSS.Range("T5:X9").ClearContents
        
        'Aging Data Table
        WS_CSS.Range("D14:R23").ClearContents
        WS_CSS.Range("D28:R28").ClearContents
        
        'Quarter Stats Table
        WS_CSS.Range("D34:W48").ClearContents
    
    Next i

    Call pCleanDBExclusive
    For j = 0 To c - 1
        Call ticketCountAll(j)
    Next j
    Call agingCountForAll
End Sub
Sub pCleanDBExclusive()
    Dim CSSlro As Integer
    'Support Dashboard contents cleansing
    'Active Ticket's stat table
    WS_CSS.Range("D5:R9").ClearContents
    WS_CSS.Range("T5:X9").ClearContents
    
    'Aging Data Table
    WS_CSS.Range("D14:R23").ClearContents
    WS_CSS.Range("D28:R28").ClearContents
    CSSlro = WS_CSS.Cells(WS_DA.Rows.Count, "C").End(xlUp).Row
    'Quarter Stats Table
    WS_CSS.Range("D34:W" & CSSlro).ClearContents
End Sub
Sub ReplicateMainSheet(ByVal Item As String)

'   Deleting the existing Team file then only it'll create new sheet
    For Each sheet In Worksheets
        If Item = sheet.Name Then
            'once the sheet name matched getting out of the loop
            Sheets(Item).Delete
            GoTo below
        End If
    Next sheet
below:
    'creating copy of Consolidated Support Stats sheet and renaming by team name
    Sheets("Consolidated Support Stats").Copy after:=Sheets("Consolidated Performance Audit")
    ActiveSheet.Name = Item

End Sub
