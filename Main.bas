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
    Public today As Date

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
    
    WS_DA.Visible = True
    
    'today is the Date of report
     today = Date

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Call InputDate

    Call pCleanDB
    
    'Quaterwise layout replication
    Call QtrReplication
    
    Call mainDataStaging
    
    'Unique list of teams
    Call CreateUniqueList
    
    Call teamsDashboard
    
    Call pCloseApp

    ' Determine how many seconds this code will take to run
    SecondsElapsed = Round(Timer - StartTime, 2)
    'Notify user in seconds
    MsgBox "Application ran successfully in " & SecondsElapsed & " seconds", vbInformation
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
    
        'getting team name from main data
        team = WS_DA.Cells(i, 22).Value
        Call agingCount(team)
        'Active Ticket Stat teamwise
        Call activeCount(team)
        'generating the dashboard for each team and quarterwise
        For j = 0 To c - 1
            Call ticketCount(team, j)
        Next j
        'calculating the median of closure rate teamwise
        Call medianClousre(team)
        
        'replicating the team's dashboard
        Call ReplicateMainSheet(team)
        'cleansing of CSS sheet after replication
        Call pCleanDBExclusive
    Next i
    
    Call pCleanDBExclusive
    
    'calling the below procedures only for Consolidated Report
    
    Call agingCountForAll
    For j = 0 To c - 1
        Call ticketCountAll(j)
    Next j
    Call activeCountAll
    Call medianClousreAll
End Sub
Sub pCleanDBExclusive()
    Dim CSSlro As Integer
    CSSlro = WS_CSS.Cells(WS_DA.Rows.Count, "C").End(xlUp).Row
    'Support Dashboard contents cleansing
    'Active Ticket's stat table
    WS_CSS.Range("D5:R9").ClearContents
    WS_CSS.Range("T5:X9").ClearContents
    
    'Aging Data Table
    WS_CSS.Range("D14:R23").ClearContents
    WS_CSS.Range("D28:R28").ClearContents
    
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
