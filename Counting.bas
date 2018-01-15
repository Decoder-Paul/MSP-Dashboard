Attribute VB_Name = "Counting"
Public Rng As Range
Sub ticketCount(ByVal team As String, ByVal v As Integer)
'========================================================================================================
' TicketCount
' -------------------------------------------------------------------------------------------------------
' Purpose   :   To get no. of Tickets depending on conditions from Data File to Dashboard .
'
' Author    :   Subhankar Paul, 11th January, 2018
' Notes     :   Different Ticket Types: 'INC', 'SRQ', 'ACT', 'PRB' are string constant
'
' Parameter :   team,v - team and quarter wise the procedure will be called repeatedly, v is the row no. of
'               quarter array
' Returns   :   N/A
' ---------------------------------------------------------------
' Revision History
'========================================================================================================
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    '----------Incident----------
    Dim opnBal_Inc(4) As Integer
    Dim recvd_Inc(4) As Integer
    Dim carried_Inc(4) As Integer
    Dim closed_Inc(4) As Integer
    Dim reOpen_Inc(4) As Integer
    Dim totEff_Inc(4) As Long
    Dim avgEff_Inc(4) As Integer
    Dim tmSize_Inc(4) As Integer
    Dim rspSla_Inc(4) As Integer
    Dim rspSlaPrcnt_Inc(4) As Integer
    Dim resSla_Inc(4) As Integer
    Dim resSlaPrcnt_Inc(4) As Integer
    Dim avgClosure_Inc(4) As Integer
    '--------Service Request--------
    Dim opnBal_Srq(4) As Integer
    Dim recvd_Srq(4) As Integer
    Dim carried_Srq(4) As Integer
    Dim closed_Srq(4) As Integer
    Dim reOpen_Srq(4) As Integer
    Dim totEff_Srq(4) As Long
    Dim avgEff_Srq(4) As Integer
    Dim tmSize_Srq(4) As Integer
    Dim rspSla_Srq(4) As Integer
    Dim rspSlaPrcnt_Srq(4) As Integer
    Dim resSla_Srq(4) As Integer
    Dim resSlaPrcnt_Srq(4) As Integer
    Dim avgClosure_Srq(4) As Integer
    '--------Problem Statement--------
    Dim opnBal_Prb(4) As Integer
    Dim recvd_Prb(4) As Integer
    Dim carried_Prb(4) As Integer
    Dim closed_Prb(4) As Integer
    Dim reOpen_Prb(4) As Integer
    Dim totEff_Prb(4) As Long
    Dim avgEff_Prb(4) As Integer
    Dim tmSize_Prb(4) As Integer
    Dim rspSla_Prb(4) As Integer
    Dim rspSlaPrcnt_Prb(4) As Integer
    Dim resSla_Prb(4) As Integer
    Dim resSlaPrcnt_Prb(4) As Integer
    Dim avgClosure_Prb(4) As Integer
    '--------Change Request--------
    Dim opnBal_Chg(4) As Integer
    Dim recvd_Chg(4) As Integer
    Dim carried_Chg(4) As Integer
    Dim closed_Chg(4) As Integer
    Dim reOpen_Chg(4) As Integer
    Dim totEff_Chg(4) As Long
    Dim avgEff_Chg(4) As Integer
    Dim tmSize_Chg(4) As Integer
    Dim winMiss_Chg(4) As Integer
    Dim winMissPrcnt_Chg(4) As Integer
    '---- variables to store the required values of each record for computation -----
    Dim Data_rowCount, Data_i, j As Long
    Dim tkt_type, rspnd, resl, person, status, reOpnd As String
    Dim prty As Integer
    Dim effort As Double
    Dim open_date As Long
    Dim closed_date As Long

    Dim age_of_tkt As Variant
    Dim startDate As Long
    Dim endDate As Long
    
    WS_DA.Select

    'parameter v is used to get the quarter version
    startDate = CLng(quarters(v, 0))
    endDate = CLng(quarters(v, 1))
    
    Data_rowCount = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row

    For Data_i = 2 To Data_rowCount
        '------------------ Filtering Data for TEAM ----------------------
        If Cells(Data_i, 8).Value = team Then
            'Opening balance: finish date = "" and create Date <= Start Date
            If Cells(Data_i, 25).Value = "" And Cells(Data_i, 23).Value <= startDate Then
                tkt_type = Cells(Data_i, 1).Value
                prty = Cells(Data_i, 12).Value
                
            '------------------ Filtering Data for Quarter ----------------------
            'Resolved: end_Date >= 'Finish Date' >= start_Date
            ElseIf Cells(Data_i, 25).Value >= startDate And Cells(Data_i, 25).Value <= endDate Then
                'Initialising the variables
                tkt_type = Cells(Data_i, 1).Value   'Ticket Type
                
                resl = Cells(Data_i, 3).Value       'Resl_SLA_Met
                prty = Cells(Data_i, 12).Value      'Priority
                effort = Cells(Data_i, 13).Value    'Actual effort(min)
                status = Cells(Data_i, 15).Value    'Status
                reOpnd = Cells(Data_i, 18).Value    'Re-opened(Y/N)
                Aging = Cells(Data_i, 19).Value     'Aging
            
            'Received: end_Date >= 'Create Date' >= start_Date
            ElseIf Cells(Data_i, 23).Value >= startDate And Cells(Data_i, 23).Value <= endDate Then
                tkt_type = Cells(Data_i, 1).Value
                prty = Cells(Data_i, 12).Value
                rspnd = Cells(Data_i, 2).Value      'Resp_SLA_Met
                
                'Carried Forward: finish date = "" and end_Date >= 'Create Date' >= start_Date
                If Cells(Data_i, 25).Value = "" Then
                    
                End If
            End If
        End If
    Next Data_i
    
    Selection.AutoFilter
    With Selection
        
        .AutoFilter Field:=8, Criteria1:=team

        .AutoFilter Field:=25, Criteria1:=">=" & CLng(quarters(v, 0)), Operator:=xlAnd, Criteria2:="<=" & CLng(quarters(v, 1))
            
    
        






    '------------------ Values Placement of the Variable in Excel sheet -------------
    '----------Incident----------
    Range("D34:H34").Value = opnBal_Inc
    Range("D35:H35").Value = recvd_Inc
    Range("D36:H36").Value = carried_Inc
    Range("D37:H37").Value = closed_Inc
    Range("D38:H38").Value = reOpen_Inc
    Range("D39:H39").Value = totEff_Inc
    Range("D40:H40").Value = avgEff_Inc
    Range("D41:H41").Value = tmSize_Inc
    Range("D44:H44").Value = rspSla_Inc
    Range("D45:H45").Value = rspSlaPrcnt_Inc
    Range("D46:H46").Value = resSla_Inc
    Range("D47:H47").Value = resSlaPrcnt_Inc
    Range("D48:H48").Value = avgClosure_Inc
    '----------Service Request----------
    Range("I34:M34").Value = opnBal_Srq
    Range("I35:M35").Value = recvd_Srq
    Range("I36:M36").Value = carried_Srq
    Range("I37:M37").Value = closed_Srq
    Range("I38:M38").Value = reOpen_Srq
    Range("I39:M39").Value = totEff_Srq
    Range("I40:M40").Value = avgEff_Srq
    Range("I41:M41").Value = tmSize_Srq
    Range("I44:M44").Value = rspSla_Srq
    Range("I45:M45").Value = rspSlaPrcnt_Srq
    Range("I46:M46").Value = resSla_Srq
    Range("I47:M47").Value = resSlaPrcnt_Srq
    Range("I48:M48").Value = avgClosure_Srq
    '----------Problem----------
    Range("N34:R34").Value = opnBal_Prb
    Range("N35:R35").Value = recvd_Prb
    Range("N36:R36").Value = carried_Prb
    Range("N37:R37").Value = closed_Prb
    Range("N38:R38").Value = reOpen_Prb
    Range("N39:R39").Value = totEff_Prb
    Range("N40:R40").Value = avgEff_Prb
    Range("N41:R41").Value = tmSize_Prb
    Range("N44:R44").Value = rspSla_Prb
    Range("N45:R45").Value = rspSlaPrcnt_Prb
    Range("N46:R46").Value = resSla_Prb
    Range("N47:R47").Value = resSlaPrcnt_Prb
    Range("N48:R48").Value = avgClosure_Prb
    '----------Change Req----------
    Range("S34:W34").Value = opnBal_Chg
    Range("S35:W35").Value = recvd_Chg
    Range("S36:W36").Value = carried_Chg
    Range("S37:W37").Value = closed_Chg
    Range("S38:W38").Value = reOpen_Chg
    Range("S39:W39").Value = totEff_Chg
    Range("S40:W40").Value = avgEff_Chg
    Range("S41:W41").Value = tmSize_Chg
    Range("S42:W42").Value = winMiss_Chg
    Range("S43:W43").Value = winMissPrcnt_Chg
End Sub

Sub Trans_SRQ_ver1()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Transformer/
    .AutoFilter Field:=9, Criteria1:="Transformers"
    ' For All SRQ Tickets
    .AutoFilter Field:=2, Criteria1:="SRQ"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        
        ' SRQ-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(4, 2).Value = Count
     '   Sheets("Summary").Cells(4, 14).Value = WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        
        ' SRQ-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(5, 2).Value = Count
        Sheets("Summary").Cells(5, 14).Value = WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        
        ' SRQ-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(6, 2).Value = Count
        Sheets("Summary").Cells(6, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))

        ' SRQ-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(7, 2).Value = Count
        Sheets("Summary").Cells(7, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' SRQ-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(10, 2).Value = Count
        Sheets("Summary").Cells(8, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' SRQ-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver1_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(8, 2).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' SRQ-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(9, 2).Value = Count
            
        .AutoFilter Field:=18
        
        ' SRQ-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver1_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver1_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(11, 2).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
