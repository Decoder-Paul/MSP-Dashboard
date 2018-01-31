Attribute VB_Name = "Counting"
Public Rng As Range
Sub ticketCount(ByVal team As String, ByVal v As Integer)
'========================================================================================================
' TicketCount
' -------------------------------------------------------------------------------------------------------
' Purpose   :   To get no. of Tickets depending on conditions from Data File to Dashboard .
'
' Author    :   Subhankar Paul, 11th January, 2018
' Notes     :   . Different Ticket Types: 'INC', 'SRQ', 'ACT', 'PRB' are string constant
'               . 4 parameters are calculated at the end of the count are
'                 Avg Effort, ResponseSLA %, ResolutionSLA % & Avg Closure Duration
'               . The value placement in Excel Cells are flexible with versions
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
    Dim totEff_Inc(4) As Double
    Dim avgEff_Inc(4) As Double
    Dim tmSize_Inc As Integer
    Dim rspSla_Inc(4) As Integer
    Dim rspSlaPrcnt_Inc(4) As Integer
    Dim resSla_Inc(4) As Integer
    Dim resSlaPrcnt_Inc(4) As Integer
    Dim cloDur_Inc(4) As Long
    Dim avgClosure_Inc(4) As Integer
    '--------Service Request--------
    Dim opnBal_Srq(4) As Integer
    Dim recvd_Srq(4) As Integer
    Dim carried_Srq(4) As Integer
    Dim closed_Srq(4) As Integer
    Dim reOpen_Srq(4) As Integer
    Dim totEff_Srq(4) As Double
    Dim avgEff_Srq(4) As Double
    Dim tmSize_Srq As Integer
    Dim rspSla_Srq(4) As Integer
    Dim rspSlaPrcnt_Srq(4) As Integer
    Dim resSla_Srq(4) As Integer
    Dim resSlaPrcnt_Srq(4) As Integer
    Dim cloDur_Srq(4) As Long
    Dim avgClosure_Srq(4) As Integer
    '--------Problem Statement--------
    Dim opnBal_Prb(4) As Integer
    Dim recvd_Prb(4) As Integer
    Dim carried_Prb(4) As Integer
    Dim closed_Prb(4) As Integer
    Dim reOpen_Prb(4) As Integer
    Dim totEff_Prb(4) As Double
    Dim avgEff_Prb(4) As Double
    Dim tmSize_Prb As Integer
    Dim rspSla_Prb(4) As Integer
    Dim rspSlaPrcnt_Prb(4) As Integer
    Dim resSla_Prb(4) As Integer
    Dim resSlaPrcnt_Prb(4) As Integer
    Dim cloDur_Prb(4) As Long
    Dim avgClosure_Prb(4) As Integer
    '--------Change Request--------
    Dim opnBal_Chg(4) As Integer
    Dim recvd_Chg(4) As Integer
    Dim carried_Chg(4) As Integer
    Dim closed_Chg(4) As Integer
    Dim reOpen_Chg(4) As Integer
    Dim totEff_Chg(4) As Double
    Dim avgEff_Chg(4) As Double
    Dim tmSize_Chg As Integer
    Dim winMiss_Chg(4) As Integer
    Dim winMissPrcnt_Chg(4) As Integer
    '---- variables to store the required values of each record for computation -----
    Dim Data_rowCount As Long
    Dim Data_i As Long
    Dim j As Long
    
    Dim tkt_type As String
    Dim reOpnd As String
    Dim rspSLA As String
    Dim resSLA As String
    Dim prty As Integer
    Dim effort As Double
    Dim createDate As Long
    Dim finishDate As Long
    '------------ Dictionary Creation for Distinct Count of Assigned resource --
    Dim INC_Dict, CHG_Dict, SRQ_Dict, PRB_Dict As Object
    Set INC_Dict = CreateObject("scripting.dictionary")
    Set CHG_Dict = CreateObject("scripting.dictionary")
    Set SRQ_Dict = CreateObject("scripting.dictionary")
    Set PRB_Dict = CreateObject("scripting.dictionary")
    
    
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
            tkt_type = Cells(Data_i, 1).Value ' Ticket Type
            prty = Cells(Data_i, 12).Value ' Priority
            createDate = Cells(Data_i, 23).Value
            finishDate = Cells(Data_i, 25).Value
            reOpnd = Cells(Data_i, 18).Value
            rspSLA = Cells(Data_i, 2).Value
            resSLA = Cells(Data_i, 3).Value
            Select Case tkt_type
                'If Incident ticket type
                Case "INC"
                    If INC_Dict.Exists(element) Then
                        INC_Dict.Item(element) = INC_Dict.Item(element) + 1
                    Else
                        INC_Dict.Add element, 1
                    End If
                Select Case prty
                    'Priority 1 and so on same for below also
                    Case 1
                        '------------------ Filtering Data for Quarter ----------------------
                        'Opening balance: finish date = "" and create Date <= Start Date
                        If createDate < startDate Then
                            'Opening Balance:
                            opnBal_Inc(0) = opnBal_Inc(0) + 1
                            'CarryForward
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Inc(0) = carried_Inc(0) + 1
                            'Resolved: end_Date >= 'Finish Date' >= start_Date
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Inc(0) = closed_Inc(0) + 1
                                'Total Effort Spent
                                totEff_Inc(0) = totEff_Inc(0) + Cells(Data_i, 13).Value
                                'Total Closure Duration
                                cloDur_Inc(0) = cloDur_Inc(0) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    'Total Resolution SLA Breached
                                    resSla_Inc(0) = resSla_Inc(0) + 1
                                End If
                            End If
                        'Received: end_Date >= 'Create Date' >= start_Date
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Inc(0) = recvd_Inc(0) + 1
                            If rspSLA = "N" Then
                                'Total Response SLA Breached
                                rspSla_Inc(0) = rspSla_Inc(0) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Inc(0) = reOpen_Inc(0) + 1
                            End If
                            'Carried Forward:
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Inc(0) = carried_Inc(0) + 1
                            'Resolved: end_Date >= 'Finish Date' >= start_Date
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Inc(0) = closed_Inc(0) + 1
                                'Total Effort Spent
                                totEff_Inc(0) = totEff_Inc(0) + Cells(Data_i, 13).Value
                                'Total Closure Duration
                                cloDur_Inc(0) = cloDur_Inc(0) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    'Total Resolution SLA Breached
                                    resSla_Inc(0) = resSla_Inc(0) + 1
                                End If
                            End If
                        End If
                    Case 2
                        If createDate < startDate Then
                            opnBal_Inc(1) = opnBal_Inc(1) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Inc(1) = carried_Inc(1) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Inc(1) = closed_Inc(1) + 1
                                totEff_Inc(1) = totEff_Inc(1) + Cells(Data_i, 13).Value
                                cloDur_Inc(1) = cloDur_Inc(1) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Inc(1) = resSla_Inc(1) + 1
                                End If
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Inc(1) = recvd_Inc(1) + 1
                            If rspSLA = "N" Then
                                rspSla_Inc(1) = rspSla_Inc(1) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Inc(1) = reOpen_Inc(1) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Inc(1) = carried_Inc(1) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Inc(1) = closed_Inc(1) + 1
                                totEff_Inc(1) = totEff_Inc(1) + Cells(Data_i, 13).Value
                                cloDur_Inc(1) = cloDur_Inc(1) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Inc(1) = resSla_Inc(1) + 1
                                End If
                            End If
                        End If
                    Case 3
                        If createDate < startDate Then
                            opnBal_Inc(2) = opnBal_Inc(2) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Inc(2) = carried_Inc(2) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Inc(2) = closed_Inc(2) + 1
                                totEff_Inc(2) = totEff_Inc(2) + Cells(Data_i, 13).Value
                                cloDur_Inc(2) = cloDur_Inc(2) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Inc(2) = resSla_Inc(2) + 1
                                End If
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Inc(2) = recvd_Inc(2) + 1
                            If rspSLA = "N" Then
                                rspSla_Inc(2) = rspSla_Inc(2) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Inc(2) = reOpen_Inc(2) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Inc(2) = carried_Inc(2) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Inc(2) = closed_Inc(2) + 1
                                totEff_Inc(2) = totEff_Inc(2) + Cells(Data_i, 13).Value
                                cloDur_Inc(2) = cloDur_Inc(2) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Inc(2) = resSla_Inc(2) + 1
                                End If
                            End If
                        End If
                    Case 4
                        If createDate < startDate Then
                            opnBal_Inc(3) = opnBal_Inc(3) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Inc(3) = carried_Inc(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Inc(3) = closed_Inc(3) + 1
                                totEff_Inc(3) = totEff_Inc(3) + Cells(Data_i, 13).Value
                                cloDur_Inc(3) = cloDur_Inc(3) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Inc(3) = resSla_Inc(3) + 1
                                End If
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Inc(3) = recvd_Inc(3) + 1
                            If rspSLA = "N" Then
                                rspSla_Inc(3) = rspSla_Inc(3) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Inc(3) = reOpen_Inc(3) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Inc(3) = carried_Inc(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Inc(3) = closed_Inc(3) + 1
                                totEff_Inc(3) = totEff_Inc(3) + Cells(Data_i, 13).Value
                                cloDur_Inc(3) = cloDur_Inc(3) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Inc(3) = resSla_Inc(3) + 1
                                End If
                            End If
                        End If
'************* Case 4 and Case 5 both sharing same variables ***************
                    Case 5
                        If createDate < startDate Then
                            opnBal_Inc(3) = opnBal_Inc(3) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Inc(3) = carried_Inc(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Inc(3) = closed_Inc(3) + 1
                                totEff_Inc(3) = totEff_Inc(3) + Cells(Data_i, 13).Value
                                cloDur_Inc(3) = cloDur_Inc(3) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Inc(3) = resSla_Inc(3) + 1
                                End If
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Inc(3) = recvd_Inc(3) + 1
                            If rspSLA = "N" Then
                                rspSla_Inc(3) = rspSla_Inc(3) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Inc(3) = reOpen_Inc(3) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Inc(3) = carried_Inc(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Inc(3) = closed_Inc(3) + 1
                                totEff_Inc(3) = totEff_Inc(3) + Cells(Data_i, 13).Value
                                cloDur_Inc(3) = cloDur_Inc(3) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Inc(3) = resSla_Inc(3) + 1
                                End If
                            End If
                        End If
                End Select
'If ticket type = SRQ
                Case "SRQ"
                    If SRQ_Dict.Exists(element) Then
                        SRQ_Dict.Item(element) = SRQ_Dict.Item(element) + 1
                    Else
                        SRQ_Dict.Add element, 1
                    End If
                Select Case prty
                    'Priority 1 and so on same for below also
                    Case 1
                        '------------------ Filtering Data for Quarter ----------------------
                        'Opening balance: finish date = "" and create Date <= Start Date
                        If createDate < startDate Then
                            'Carried Forward & Opening Balance:
                            opnBal_Srq(0) = opnBal_Srq(0) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Srq(0) = carried_Srq(0) + 1
                            'Resolved: end_Date >= 'Finish Date' >= start_Date
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Srq(0) = closed_Srq(0) + 1
                                'Total Effort Spent
                                totEff_Srq(0) = totEff_Srq(0) + Cells(Data_i, 13).Value
                                'Total Closure Duration
                                cloDur_Srq(0) = cloDur_Srq(0) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    'Total Resolution SLA Breached
                                    resSla_Srq(0) = resSla_Srq(0) + 1
                                End If
                            End If
                        'Received: end_Date >= 'Create Date' >= start_Date
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Srq(0) = recvd_Srq(0) + 1
                            If rspSLA = "N" Then
                                'Total Response SLA Breached
                                rspSla_Srq(0) = rspSla_Srq(0) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Srq(0) = reOpen_Srq(0) + 1
                            End If
                            'Carried Forward:
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Srq(0) = carried_Srq(0) + 1
                            'Resolved: end_Date >= 'Finish Date' >= start_Date
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Srq(0) = closed_Srq(0) + 1
                                'Total Effort Spent
                                totEff_Srq(0) = totEff_Srq(0) + Cells(Data_i, 13).Value
                                'Total Closure Duration
                                cloDur_Srq(0) = cloDur_Srq(0) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    'Total Resolution SLA Breached
                                    resSla_Srq(0) = resSla_Srq(0) + 1
                                End If
                            End If
                        End If
                    Case 2
                        If createDate < startDate Then
                            opnBal_Srq(1) = opnBal_Srq(1) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Srq(1) = carried_Srq(1) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Srq(1) = closed_Srq(1) + 1
                                totEff_Srq(1) = totEff_Srq(1) + Cells(Data_i, 13).Value
                                cloDur_Srq(1) = cloDur_Srq(1) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Srq(1) = resSla_Srq(1) + 1
                                End If
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Srq(1) = recvd_Srq(1) + 1
                            If rspSLA = "N" Then
                                rspSla_Srq(1) = rspSla_Srq(1) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Srq(1) = reOpen_Srq(1) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Srq(1) = carried_Srq(1) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Srq(1) = closed_Srq(1) + 1
                                totEff_Srq(1) = totEff_Srq(1) + Cells(Data_i, 13).Value
                                cloDur_Srq(1) = cloDur_Srq(1) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Srq(1) = resSla_Srq(1) + 1
                                End If
                            End If
                        End If
                    Case 3
                        If createDate < startDate Then
                            opnBal_Srq(2) = opnBal_Srq(2) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Srq(2) = carried_Srq(2) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Srq(2) = closed_Srq(2) + 1
                                totEff_Srq(2) = totEff_Srq(2) + Cells(Data_i, 13).Value
                                cloDur_Srq(2) = cloDur_Srq(2) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Srq(2) = resSla_Srq(2) + 1
                                End If
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Srq(2) = recvd_Srq(2) + 1
                            If rspSLA = "N" Then
                                rspSla_Srq(2) = rspSla_Srq(2) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Srq(2) = reOpen_Srq(2) + 1
                            End If
                            If CStr(finishDate) = o Or finishDate > endDate Then
                                carried_Srq(2) = carried_Srq(2) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Srq(2) = closed_Srq(2) + 1
                                totEff_Srq(2) = totEff_Srq(2) + Cells(Data_i, 13).Value
                                cloDur_Srq(2) = cloDur_Srq(2) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Srq(2) = resSla_Srq(2) + 1
                                End If
                            End If
                        End If
                    Case 4
                        If createDate < startDate Then
                            opnBal_Srq(3) = opnBal_Srq(3) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Srq(3) = carried_Srq(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Srq(3) = closed_Srq(3) + 1
                                totEff_Srq(3) = totEff_Srq(3) + Cells(Data_i, 13).Value
                                cloDur_Srq(3) = cloDur_Srq(3) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Srq(3) = resSla_Srq(3) + 1
                                End If
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Srq(3) = recvd_Srq(3) + 1
                            If rspSLA = "N" Then
                                rspSla_Srq(3) = rspSla_Srq(3) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Srq(3) = reOpen_Srq(3) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Srq(3) = carried_Srq(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Srq(3) = closed_Srq(3) + 1
                                totEff_Srq(3) = totEff_Srq(3) + Cells(Data_i, 13).Value
                                cloDur_Srq(3) = cloDur_Srq(3) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Srq(3) = resSla_Srq(3) + 1
                                End If
                            End If
                        End If
'************* Case 4 and Case 5 both sharing same variables ***************
                    Case 5
                        If createDate < startDate Then
                            opnBal_Srq(3) = opnBal_Srq(3) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Srq(3) = carried_Srq(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Srq(3) = closed_Srq(3) + 1
                                totEff_Srq(3) = totEff_Srq(3) + Cells(Data_i, 13).Value
                                cloDur_Srq(3) = cloDur_Srq(3) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Srq(3) = resSla_Srq(3) + 1
                                End If
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Srq(3) = recvd_Srq(3) + 1
                            If rspSLA = "N" Then
                                rspSla_Srq(3) = rspSla_Srq(3) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Srq(3) = reOpen_Srq(3) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Srq(3) = carried_Srq(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Srq(3) = closed_Srq(3) + 1
                                totEff_Srq(3) = totEff_Srq(3) + Cells(Data_i, 13).Value
                                cloDur_Srq(3) = cloDur_Srq(3) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Srq(3) = resSla_Srq(3) + 1
                                End If
                            End If
                        End If
                    End Select
'If ticket type = PRB
                Case "PRB"
                    If PRB_Dict.Exists(element) Then
                        PRB_Dict.Item(element) = PRB_Dict.Item(element) + 1
                    Else
                        PRB_Dict.Add element, 1
                    End If
                Select Case prty
                    'Priority 1 and so on same for below also
                    Case 1
                        '------------------ Filtering Data for Quarter ----------------------
                        'Opening balance: finish date = "" and create Date <= Start Date
                        If createDate < startDate Then
                            'Carried Forward & Opening Balance:
                            opnBal_Prb(0) = opnBal_Prb(0) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Prb(0) = carried_Prb(0) + 1
                            'Resolved: end_Date >= 'Finish Date' >= start_Date
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Prb(0) = closed_Prb(0) + 1
                                'Total Effort Spent
                                totEff_Prb(0) = totEff_Prb(0) + Cells(Data_i, 13).Value
                                'Total Closure Duration
                                cloDur_Prb(0) = cloDur_Prb(0) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    'Total Resolution SLA Breached
                                    resSla_Prb(0) = resSla_Prb(0) + 1
                                End If
                            End If
                        'Received: end_Date >= 'Create Date' >= start_Date
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Prb(0) = recvd_Prb(0) + 1
                            If rspSLA = "N" Then
                                'Total Response SLA Breached
                                rspSla_Prb(0) = rspSla_Prb(0) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Prb(0) = reOpen_Prb(0) + 1
                            End If
                            'Carried Forward:
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Prb(0) = carried_Prb(0) + 1
                            'Resolved: end_Date >= 'Finish Date' >= start_Date
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Prb(0) = closed_Prb(0) + 1
                                'Total Effort Spent
                                totEff_Prb(0) = totEff_Prb(0) + Cells(Data_i, 13).Value
                                'Total Closure Duration
                                cloDur_Prb(0) = cloDur_Prb(0) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    'Total Resolution SLA Breached
                                    resSla_Prb(0) = resSla_Prb(0) + 1
                                End If
                            End If
                        End If
                    Case 2
                        If createDate < startDate Then
                            opnBal_Prb(1) = opnBal_Prb(1) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Prb(1) = carried_Prb(1) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Prb(1) = closed_Prb(1) + 1
                                totEff_Prb(1) = totEff_Prb(1) + Cells(Data_i, 13).Value
                                cloDur_Prb(1) = cloDur_Prb(1) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Prb(1) = resSla_Prb(1) + 1
                                End If
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Prb(1) = recvd_Prb(1) + 1
                            If rspSLA = "N" Then
                                rspSla_Prb(1) = rspSla_Prb(1) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Prb(1) = reOpen_Prb(1) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Prb(1) = carried_Prb(1) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Prb(1) = closed_Prb(1) + 1
                                totEff_Prb(1) = totEff_Prb(1) + Cells(Data_i, 13).Value
                                cloDur_Prb(1) = cloDur_Prb(1) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Prb(1) = resSla_Prb(1) + 1
                                End If
                            End If
                        End If
                    Case 3
                        If createDate < startDate Then
                            opnBal_Prb(2) = opnBal_Prb(2) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Prb(2) = carried_Prb(2) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Prb(2) = closed_Prb(2) + 1
                                totEff_Prb(2) = totEff_Prb(2) + Cells(Data_i, 13).Value
                                cloDur_Prb(2) = cloDur_Prb(2) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Prb(2) = resSla_Prb(2) + 1
                                End If
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Prb(2) = recvd_Prb(2) + 1
                            If rspSLA = "N" Then
                                rspSla_Prb(2) = rspSla_Prb(2) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Prb(2) = reOpen_Prb(2) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Prb(2) = carried_Prb(2) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Prb(2) = closed_Prb(2) + 1
                                totEff_Prb(2) = totEff_Prb(2) + Cells(Data_i, 13).Value
                                cloDur_Prb(2) = cloDur_Prb(2) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Prb(2) = resSla_Prb(2) + 1
                                End If
                            End If
                        End If
                    Case 4
                        If createDate < startDate Then
                            opnBal_Prb(3) = opnBal_Prb(3) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Prb(3) = carried_Prb(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Prb(3) = closed_Prb(3) + 1
                                totEff_Prb(3) = totEff_Prb(3) + Cells(Data_i, 13).Value
                                cloDur_Prb(3) = cloDur_Prb(3) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Prb(3) = resSla_Prb(3) + 1
                                End If
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Prb(3) = recvd_Prb(3) + 1
                            If rspSLA = "N" Then
                                rspSla_Prb(3) = rspSla_Prb(3) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Prb(3) = reOpen_Prb(3) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Prb(3) = carried_Prb(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Prb(3) = closed_Prb(3) + 1
                                totEff_Prb(3) = totEff_Prb(3) + Cells(Data_i, 13).Value
                                cloDur_Prb(3) = cloDur_Prb(3) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Prb(3) = resSla_Prb(3) + 1
                                End If
                            End If
                        End If
'************* Case 4 and Case 5 both sharing same variables ***************
                    Case 5
                        If createDate < startDate Then
                            opnBal_Prb(3) = opnBal_Prb(3) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Prb(3) = carried_Prb(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Prb(3) = closed_Prb(3) + 1
                                totEff_Prb(3) = totEff_Prb(3) + Cells(Data_i, 13).Value
                                cloDur_Prb(3) = cloDur_Prb(3) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Prb(3) = resSla_Prb(3) + 1
                                End If
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Prb(3) = recvd_Prb(3) + 1
                            If rspSLA = "N" Then
                                rspSla_Prb(3) = rspSla_Prb(3) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Prb(3) = reOpen_Prb(3) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Prb(3) = carried_Prb(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Prb(3) = closed_Prb(3) + 1
                                totEff_Prb(3) = totEff_Prb(3) + Cells(Data_i, 13).Value
                                cloDur_Prb(3) = cloDur_Prb(3) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Prb(3) = resSla_Prb(3) + 1
                                End If
                            End If
                        End If
                End Select
                'If ticket type = CHG
                Case "ACT"
                    If CHG_Dict.Exists(element) Then
                        CHG_Dict.Item(element) = CHG_Dict.Item(element) + 1
                    Else
                        CHG_Dict.Add element, 1
                    End If
                Select Case prty
                    'Priority 1 and so on same for below also
                    Case 1
                        '------------------ Filtering Data for Quarter ----------------------
                        'Opening balance: finish date = "" and create Date <= Start Date
                        If createDate < startDate Then
                            'Carried Forward & Opening Balance:
                            opnBal_Chg(0) = opnBal_Chg(0) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Chg(0) = carried_Chg(0) + 1
                            'Resolved: end_Date >= 'Finish Date' >= start_Date
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Chg(0) = closed_Chg(0) + 1
                                'Total Effort Spent
                                totEff_Chg(0) = totEff_Chg(0) + Cells(Data_i, 13).Value
                            End If
                        'Received: end_Date >= 'Create Date' >= start_Date
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Chg(0) = recvd_Chg(0) + 1
                            If reOpnd = "Y" Then
                                reOpen_Chg(0) = reOpen_Chg(0) + 1
                            End If
                            'Carried Forward:
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Chg(0) = carried_Chg(0) + 1
                            'Resolved: end_Date >= 'Finish Date' >= start_Date
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Chg(0) = closed_Chg(0) + 1
                                'Total Effort Spent
                                totEff_Chg(0) = totEff_Chg(0) + Cells(Data_i, 13).Value
                                'Change Window missed implemented
                                If finishDate < Cells(Data_i, 16).Value And finishDate > Cells(Data_i, 17).Value Then
                                    winMiss_Chg(0) = winMiss_Chg(0) + 1
                                End If
                            End If
                        End If
                    Case 2
                        If createDate < startDate Then
                            opnBal_Chg(1) = opnBal_Chg(1) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Chg(1) = carried_Chg(1) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Chg(1) = closed_Chg(1) + 1
                                totEff_Chg(1) = totEff_Chg(1) + Cells(Data_i, 13).Value
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Chg(1) = recvd_Chg(1) + 1
                            If reOpnd = "Y" Then
                                reOpen_Chg(1) = reOpen_Chg(1) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Chg(1) = carried_Chg(1) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Chg(1) = closed_Chg(1) + 1
                                totEff_Chg(1) = totEff_Chg(1) + Cells(Data_i, 13).Value
                                If finishDate < Cells(Data_i, 16).Value And finishDate > Cells(Data_i, 17).Value Then
                                    winMiss_Chg(1) = winMiss_Chg(1) + 1
                                End If
                            End If
                        End If
                    Case 3
                        If createDate < startDate Then
                            opnBal_Chg(2) = opnBal_Chg(2) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Chg(2) = carried_Chg(2) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Chg(2) = closed_Chg(2) + 1
                                totEff_Chg(2) = totEff_Chg(2) + Cells(Data_i, 13).Value
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Chg(2) = recvd_Chg(2) + 1
                            If reOpnd = "Y" Then
                                reOpen_Chg(2) = reOpen_Chg(2) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Chg(2) = carried_Chg(2) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Chg(2) = closed_Chg(2) + 1
                                totEff_Chg(2) = totEff_Chg(2) + Cells(Data_i, 13).Value
                                If finishDate < Cells(Data_i, 16).Value And finishDate > Cells(Data_i, 17).Value Then
                                    winMiss_Chg(2) = winMiss_Chg(2) + 1
                                End If
                            End If
                        End If
                    Case 4
                        If createDate < startDate Then
                            opnBal_Chg(3) = opnBal_Chg(3) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Chg(3) = carried_Chg(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Chg(3) = closed_Chg(3) + 1
                                totEff_Chg(3) = totEff_Chg(3) + Cells(Data_i, 13).Value
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Chg(3) = recvd_Chg(3) + 1
                            If reOpnd = "Y" Then
                                reOpen_Chg(3) = reOpen_Chg(3) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Chg(3) = carried_Chg(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Chg(3) = closed_Chg(3) + 1
                                totEff_Chg(3) = totEff_Chg(3) + Cells(Data_i, 13).Value
                                If finishDate < Cells(Data_i, 16).Value And finishDate > Cells(Data_i, 17).Value Then
                                    winMiss_Chg(3) = winMiss_Chg(0) + 1
                                End If
                            End If
                        End If
'************* Case 4 and Case 5 both sharing same variables ***************
                    Case 5
                        If createDate < startDate Then
                            opnBal_Chg(3) = opnBal_Chg(3) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Chg(3) = carried_Chg(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Chg(3) = closed_Chg(3) + 1
                                totEff_Chg(3) = totEff_Chg(3) + Cells(Data_i, 13).Value
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Chg(3) = recvd_Chg(3) + 1
                            If reOpnd = "Y" Then
                                reOpen_Chg(3) = reOpen_Chg(3) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Chg(3) = carried_Chg(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Chg(3) = closed_Chg(3) + 1
                                totEff_Chg(3) = totEff_Chg(3) + Cells(Data_i, 13).Value
                                If finishDate < Cells(Data_i, 16).Value And finishDate > Cells(Data_i, 17).Value Then
                                    winMiss_Chg(3) = winMiss_Chg(3) + 1
                                End If
                            End If
                        End If
                End Select
            End Select
        End If
    Next Data_i
    ' Total of the following variables are calculated
    For i = 0 To 3
        opnBal_Inc(4) = opnBal_Inc(4) + opnBal_Inc(i)
        recvd_Inc(4) = recvd_Inc(4) + recvd_Inc(i)
        carried_Inc(4) = carried_Inc(4) + carried_Inc(i)
        closed_Inc(4) = closed_Inc(4) + closed_Inc(i)
        reOpen_Inc(4) = reOpen_Inc(4) + reOpen_Inc(i)
        totEff_Inc(4) = totEff_Inc(4) + totEff_Inc(i)
        rspSla_Inc(4) = rspSla_Inc(4) + rspSla_Inc(i)
        resSla_Inc(4) = resSla_Inc(4) + resSla_Inc(i)
        cloDur_Inc(4) = cloDur_Inc(4) + cloDur_Inc(i)
        
        opnBal_Srq(4) = opnBal_Srq(4) + opnBal_Srq(i)
        recvd_Srq(4) = recvd_Srq(4) + recvd_Srq(i)
        carried_Srq(4) = carried_Srq(4) + carried_Srq(i)
        closed_Srq(4) = closed_Srq(4) + closed_Srq(i)
        reOpen_Srq(4) = reOpen_Srq(4) + reOpen_Srq(i)
        totEff_Srq(4) = totEff_Srq(4) + totEff_Srq(i)
        rspSla_Srq(4) = rspSla_Srq(4) + rspSla_Srq(i)
        resSla_Srq(4) = resSla_Srq(4) + resSla_Srq(i)
        cloDur_Srq(4) = cloDur_Srq(4) + cloDur_Srq(i)
        
        opnBal_Prb(4) = opnBal_Prb(4) + opnBal_Prb(i)
        recvd_Prb(4) = recvd_Prb(4) + recvd_Prb(i)
        carried_Prb(4) = carried_Prb(4) + carried_Prb(i)
        closed_Prb(4) = closed_Prb(4) + closed_Prb(i)
        reOpen_Prb(4) = reOpen_Prb(4) + reOpen_Prb(i)
        totEff_Prb(4) = totEff_Prb(4) + totEff_Prb(i)
        rspSla_Prb(4) = rspSla_Prb(4) + rspSla_Prb(i)
        resSla_Prb(4) = resSla_Prb(4) + resSla_Prb(i)
        cloDur_Prb(4) = cloDur_Prb(4) + cloDur_Prb(i)
        
        opnBal_Chg(4) = opnBal_Chg(4) + opnBal_Chg(i)
        recvd_Chg(4) = recvd_Chg(4) + recvd_Chg(i)
        carried_Chg(4) = carried_Chg(4) + carried_Chg(i)
        closed_Chg(4) = closed_Chg(4) + closed_Chg(i)
        reOpen_Chg(4) = reOpen_Chg(4) + reOpen_Chg(i)
        totEff_Chg(4) = totEff_Chg(4) + totEff_Chg(i)
        winMiss_Chg(4) = winMiss_Chg(4) + winMiss_Chg(i)
    Next i
    'Conversion of Minutes to Hours for Total Effort
    For i = 0 To 4
        totEff_Inc(i) = Round(totEff_Inc(i) / 60, 2)
        totEff_Srq(i) = Round(totEff_Srq(i) / 60, 2)
        totEff_Prb(i) = Round(totEff_Prb(i) / 60, 2)
        totEff_Chg(i) = Round(totEff_Chg(i) / 60, 2)
    Next i
    'average is being calculated here
    For i = 0 To 4
        If closed_Inc(i) <> 0 Then
            avgEff_Inc(i) = Round(totEff_Inc(i) / closed_Inc(i), 2)
            resSlaPrcnt_Inc(i) = resSla_Inc(i) * 100 / closed_Inc(i)
            avgClosure_Inc(i) = cloDur_Inc(i) / closed_Inc(i)
        End If
        If closed_Srq(i) <> 0 Then
            avgEff_Srq(i) = Round(totEff_Srq(i) / closed_Srq(i), 2)
            resSlaPrcnt_Srq(i) = resSla_Srq(i) * 100 / closed_Srq(i)
            avgClosure_Srq(i) = cloDur_Srq(i) / closed_Srq(i)
        End If
        If closed_Prb(i) <> 0 Then
            avgEff_Prb(i) = Round(totEff_Prb(i) / closed_Prb(i), 2)
            resSlaPrcnt_Prb(i) = resSla_Prb(i) * 100 / closed_Prb(i)
            avgClosure_Prb(i) = cloDur_Prb(i) / closed_Prb(i)
        End If
        If closed_Chg(i) <> 0 Then
            avgEff_Chg(i) = Round(totEff_Chg(i) / closed_Chg(i), 2)
            winMissPrcnt_Chg(i) = winMiss_Chg(i) * 100 / closed_Chg(i)
        End If
        
        If recvd_Inc(i) <> 0 Then
            rspSlaPrcnt_Inc(i) = rspSla_Inc(i) * 100 / recvd_Inc(i)
        End If
        If recvd_Srq(i) <> 0 Then
            rspSlaPrcnt_Srq(i) = rspSla_Srq(i) * 100 / recvd_Srq(i)
        End If
        If recvd_Prb(i) <> 0 Then
            rspSlaPrcnt_Prb(i) = rspSla_Prb(i) * 100 / recvd_Prb(i)
        End If
    Next i
    
    'Checking Whether Dictionary contains Blank String
    If INC_Dict.Exists("") Then
        tmSize_Inc = INC_Dict.Count - 1
    End If
    If SRQ_Dict.Exists("") Then
        tmSize_Srq = SRQ_Dict.Count - 1
    End If
    If CHG_Dict.Exists("") Then
        tmSize_Chg = CHG_Dict.Count - 1
    End If
    If PRB_Dict.Exists("") Then
        tmSize_Prb = PRB_Dict.Count - 1
    End If
    
    
    WS_CSS.Select
    
    '------------------ VERSIONWISE Value Placement of the Variable in Excel sheet -------------
    '----------Incident----------
    Range("D" & 34 + 15 * v & ":H" & 34 + 15 * v).Value = opnBal_Inc
    Range("D" & 35 + 15 * v & ":H" & 35 + 15 * v).Value = recvd_Inc
    Range("D" & 36 + 15 * v & ":H" & 36 + 15 * v).Value = carried_Inc
    Range("D" & 37 + 15 * v & ":H" & 37 + 15 * v).Value = closed_Inc
    Range("D" & 38 + 15 * v & ":H" & 38 + 15 * v).Value = reOpen_Inc
    Range("D" & 39 + 15 * v & ":H" & 39 + 15 * v).Value = totEff_Inc
    Range("D" & 40 + 15 * v & ":H" & 40 + 15 * v).Value = avgEff_Inc
    Cells(41 + 15 * v, 4).Value = tmSize_Inc
    Range("D" & 44 + 15 * v & ":H" & 44 + 15 * v).Value = rspSla_Inc
    Range("D" & 45 + 15 * v & ":H" & 45 + 15 * v).Value = rspSlaPrcnt_Inc
    Range("D" & 46 + 15 * v & ":H" & 46 + 15 * v).Value = resSla_Inc
    Range("D" & 47 + 15 * v & ":H" & 47 + 15 * v).Value = resSlaPrcnt_Inc
    Range("D" & 48 + 15 * v & ":H" & 48 + 15 * v).Value = avgClosure_Inc
    '----------Service Request----------
    Range("I" & 34 + 15 * v & ":M" & 34 + 15 * v).Value = opnBal_Srq
    Range("I" & 35 + 15 * v & ":M" & 35 + 15 * v).Value = recvd_Srq
    Range("I" & 36 + 15 * v & ":M" & 36 + 15 * v).Value = carried_Srq
    Range("I" & 37 + 15 * v & ":M" & 37 + 15 * v).Value = closed_Srq
    Range("I" & 38 + 15 * v & ":M" & 38 + 15 * v).Value = reOpen_Srq
    Range("I" & 39 + 15 * v & ":M" & 39 + 15 * v).Value = totEff_Srq
    Range("I" & 40 + 15 * v & ":M" & 40 + 15 * v).Value = avgEff_Srq
    Cells(41 + 15 * v, 9).Value = tmSize_Srq
    Range("I" & 44 + 15 * v & ":M" & 44 + 15 * v).Value = rspSla_Srq
    Range("I" & 45 + 15 * v & ":M" & 45 + 15 * v).Value = rspSlaPrcnt_Srq
    Range("I" & 46 + 15 * v & ":M" & 46 + 15 * v).Value = resSla_Srq
    Range("I" & 47 + 15 * v & ":M" & 47 + 15 * v).Value = resSlaPrcnt_Srq
    Range("I" & 48 + 15 * v & ":M" & 48 + 15 * v).Value = avgClosure_Srq
    '----------Problem----------
    Range("N" & 34 + 15 * v & ":R" & 34 + 15 * v).Value = opnBal_Prb
    Range("N" & 35 + 15 * v & ":R" & 35 + 15 * v).Value = recvd_Prb
    Range("N" & 36 + 15 * v & ":R" & 36 + 15 * v).Value = carried_Prb
    Range("N" & 37 + 15 * v & ":R" & 37 + 15 * v).Value = closed_Prb
    Range("N" & 38 + 15 * v & ":R" & 38 + 15 * v).Value = reOpen_Prb
    Range("N" & 39 + 15 * v & ":R" & 39 + 15 * v).Value = totEff_Prb
    Range("N" & 40 + 15 * v & ":R" & 40 + 15 * v).Value = avgEff_Prb
    Cells(41 + 15 * v, 14).Value = tmSize_Prb
    Range("N" & 44 + 15 * v & ":R" & 44 + 15 * v).Value = rspSla_Prb
    Range("N" & 45 + 15 * v & ":R" & 45 + 15 * v).Value = rspSlaPrcnt_Prb
    Range("N" & 46 + 15 * v & ":R" & 46 + 15 * v).Value = resSla_Prb
    Range("N" & 47 + 15 * v & ":R" & 47 + 15 * v).Value = resSlaPrcnt_Prb
    Range("N" & 48 + 15 * v & ":R" & 48 + 15 * v).Value = avgClosure_Prb
    '----------Change Req----------
    Range("S" & 34 + 15 * v & ":W" & 34 + 15 * v).Value = opnBal_Chg
    Range("S" & 35 + 15 * v & ":W" & 35 + 15 * v).Value = recvd_Chg
    Range("S" & 36 + 15 * v & ":W" & 36 + 15 * v).Value = carried_Chg
    Range("S" & 37 + 15 * v & ":W" & 37 + 15 * v).Value = closed_Chg
    Range("S" & 38 + 15 * v & ":W" & 38 + 15 * v).Value = reOpen_Chg
    Range("S" & 39 + 15 * v & ":W" & 39 + 15 * v).Value = totEff_Chg
    Range("S" & 40 + 15 * v & ":W" & 40 + 15 * v).Value = avgEff_Chg
    Cells(41 + 15 * v, 19).Value = tmSize_Chg
    Range("S" & 42 + 15 * v & ":W" & 42 + 15 * v).Value = winMiss_Chg
    Range("S" & 43 + 15 * v & ":W" & 43 + 15 * v).Value = winMissPrcnt_Chg
    'Quarter Range of date placing
    Cells(34 + 15 * v, 2).Value = quarters(v, 0) & "  -  " & quarters(v, 1)
    
End Sub
Sub ticketCountAll(ByVal v As Integer)
'========================================================================================================
' TicketCountAll
' -------------------------------------------------------------------------------------------------------
' Purpose   :   To get no. of Tickets depending on conditions from Data File to Dashboard .
'
' Author    :   Subhankar Paul, 11th January, 2018
' Notes     :   . Different Ticket Types: 'INC', 'SRQ', 'ACT', 'PRB' are string constant
'               . 4 parameters are calculated at the end of the count are
'                 Avg Effort, ResponseSLA %, ResolutionSLA % & Avg Closure Duration
'               . The value placement in Excel Cells are flexible with versions
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
    Dim totEff_Inc(4) As Double
    Dim avgEff_Inc(4) As Double
    Dim tmSize_Inc As Integer
    Dim rspSla_Inc(4) As Integer
    Dim rspSlaPrcnt_Inc(4) As Integer
    Dim resSla_Inc(4) As Integer
    Dim resSlaPrcnt_Inc(4) As Integer
    Dim cloDur_Inc(4) As Long
    Dim avgClosure_Inc(4) As Integer
    '--------Service Request--------
    Dim opnBal_Srq(4) As Integer
    Dim recvd_Srq(4) As Integer
    Dim carried_Srq(4) As Integer
    Dim closed_Srq(4) As Integer
    Dim reOpen_Srq(4) As Integer
    Dim totEff_Srq(4) As Double
    Dim avgEff_Srq(4) As Double
    Dim tmSize_Srq As Integer
    Dim rspSla_Srq(4) As Integer
    Dim rspSlaPrcnt_Srq(4) As Integer
    Dim resSla_Srq(4) As Integer
    Dim resSlaPrcnt_Srq(4) As Integer
    Dim cloDur_Srq(4) As Long
    Dim avgClosure_Srq(4) As Integer
    '--------Problem Statement--------
    Dim opnBal_Prb(4) As Integer
    Dim recvd_Prb(4) As Integer
    Dim carried_Prb(4) As Integer
    Dim closed_Prb(4) As Integer
    Dim reOpen_Prb(4) As Integer
    Dim totEff_Prb(4) As Double
    Dim avgEff_Prb(4) As Double
    Dim tmSize_Prb As Integer
    Dim rspSla_Prb(4) As Integer
    Dim rspSlaPrcnt_Prb(4) As Integer
    Dim resSla_Prb(4) As Integer
    Dim resSlaPrcnt_Prb(4) As Integer
    Dim cloDur_Prb(4) As Long
    Dim avgClosure_Prb(4) As Integer
    '--------Change Request--------
    Dim opnBal_Chg(4) As Integer
    Dim recvd_Chg(4) As Integer
    Dim carried_Chg(4) As Integer
    Dim closed_Chg(4) As Integer
    Dim reOpen_Chg(4) As Integer
    Dim totEff_Chg(4) As Double
    Dim avgEff_Chg(4) As Double
    Dim tmSize_Chg As Integer
    Dim winMiss_Chg(4) As Integer
    Dim winMissPrcnt_Chg(4) As Integer
    '---- variables to store the required values of each record for computation -----
    Dim Data_rowCount As Long
    Dim Data_i As Long
    Dim j As Long
    
    Dim tkt_type As String
    Dim reOpnd As String
    Dim rspSLA As String
    Dim resSLA As String
    Dim prty As Integer
    Dim effort As Double
    Dim createDate As Long
    Dim finishDate As Long
    '------------ Dictionary Creation for Distinct Count of Assigned resource --
    Dim INC_Dict, CHG_Dict, SRQ_Dict, PRB_Dict As Object
    Set INC_Dict = CreateObject("scripting.dictionary")
    Set CHG_Dict = CreateObject("scripting.dictionary")
    Set SRQ_Dict = CreateObject("scripting.dictionary")
    Set PRB_Dict = CreateObject("scripting.dictionary")
    
    
    Dim startDate As Long
    Dim endDate As Long
    
    WS_DA.Select

    'parameter v is used to get the quarter version
    startDate = CLng(quarters(v, 0))
    endDate = CLng(quarters(v, 1))
    
    Data_rowCount = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row

    For Data_i = 2 To Data_rowCount
            tkt_type = Cells(Data_i, 1).Value ' Ticket Type
            prty = Cells(Data_i, 12).Value ' Priority
            createDate = Cells(Data_i, 23).Value
            finishDate = Cells(Data_i, 25).Value
            reOpnd = Cells(Data_i, 18).Value
            rspSLA = Cells(Data_i, 2).Value
            resSLA = Cells(Data_i, 3).Value
            Select Case tkt_type
                'If Incident ticket type
                Case "INC"
                    If INC_Dict.Exists(element) Then
                        INC_Dict.Item(element) = INC_Dict.Item(element) + 1
                    Else
                        INC_Dict.Add element, 1
                    End If
                Select Case prty
                    'Priority 1 and so on same for below also
                    Case 1
                        '------------------ Filtering Data for Quarter ----------------------
                        'Opening balance: finish date = "" and create Date <= Start Date
                        If createDate < startDate Then
                            'Opening Balance:
                            opnBal_Inc(0) = opnBal_Inc(0) + 1
                            'CarryForward
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Inc(0) = carried_Inc(0) + 1
                            'Resolved: end_Date >= 'Finish Date' >= start_Date
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Inc(0) = closed_Inc(0) + 1
                                'Total Effort Spent
                                totEff_Inc(0) = totEff_Inc(0) + Cells(Data_i, 13).Value
                                'Total Closure Duration
                                cloDur_Inc(0) = cloDur_Inc(0) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    'Total Resolution SLA Breached
                                    resSla_Inc(0) = resSla_Inc(0) + 1
                                End If
                            End If
                        'Received: end_Date >= 'Create Date' >= start_Date
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Inc(0) = recvd_Inc(0) + 1
                            If rspSLA = "N" Then
                                'Total Response SLA Breached
                                rspSla_Inc(0) = rspSla_Inc(0) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Inc(0) = reOpen_Inc(0) + 1
                            End If
                            'Carried Forward:
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Inc(0) = carried_Inc(0) + 1
                            'Resolved: end_Date >= 'Finish Date' >= start_Date
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Inc(0) = closed_Inc(0) + 1
                                'Total Effort Spent
                                totEff_Inc(0) = totEff_Inc(0) + Cells(Data_i, 13).Value
                                'Total Closure Duration
                                cloDur_Inc(0) = cloDur_Inc(0) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    'Total Resolution SLA Breached
                                    resSla_Inc(0) = resSla_Inc(0) + 1
                                End If
                            End If
                        End If
                    Case 2
                        If createDate < startDate Then
                            opnBal_Inc(1) = opnBal_Inc(1) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Inc(1) = carried_Inc(1) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Inc(1) = closed_Inc(1) + 1
                                totEff_Inc(1) = totEff_Inc(1) + Cells(Data_i, 13).Value
                                cloDur_Inc(1) = cloDur_Inc(1) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Inc(1) = resSla_Inc(1) + 1
                                End If
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Inc(1) = recvd_Inc(1) + 1
                            If rspSLA = "N" Then
                                rspSla_Inc(1) = rspSla_Inc(1) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Inc(1) = reOpen_Inc(1) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Inc(1) = carried_Inc(1) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Inc(1) = closed_Inc(1) + 1
                                totEff_Inc(1) = totEff_Inc(1) + Cells(Data_i, 13).Value
                                cloDur_Inc(1) = cloDur_Inc(1) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Inc(1) = resSla_Inc(1) + 1
                                End If
                            End If
                        End If
                    Case 3
                        If createDate < startDate Then
                            opnBal_Inc(2) = opnBal_Inc(2) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Inc(2) = carried_Inc(2) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Inc(2) = closed_Inc(2) + 1
                                totEff_Inc(2) = totEff_Inc(2) + Cells(Data_i, 13).Value
                                cloDur_Inc(2) = cloDur_Inc(2) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Inc(2) = resSla_Inc(2) + 1
                                End If
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Inc(2) = recvd_Inc(2) + 1
                            If rspSLA = "N" Then
                                rspSla_Inc(2) = rspSla_Inc(2) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Inc(2) = reOpen_Inc(2) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Inc(2) = carried_Inc(2) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Inc(2) = closed_Inc(2) + 1
                                totEff_Inc(2) = totEff_Inc(2) + Cells(Data_i, 13).Value
                                cloDur_Inc(2) = cloDur_Inc(2) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Inc(2) = resSla_Inc(2) + 1
                                End If
                            End If
                        End If
                    Case 4
                        If createDate < startDate Then
                            opnBal_Inc(3) = opnBal_Inc(3) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Inc(3) = carried_Inc(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Inc(3) = closed_Inc(3) + 1
                                totEff_Inc(3) = totEff_Inc(3) + Cells(Data_i, 13).Value
                                cloDur_Inc(3) = cloDur_Inc(3) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Inc(3) = resSla_Inc(3) + 1
                                End If
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Inc(3) = recvd_Inc(3) + 1
                            If rspSLA = "N" Then
                                rspSla_Inc(3) = rspSla_Inc(3) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Inc(3) = reOpen_Inc(3) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Inc(3) = carried_Inc(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Inc(3) = closed_Inc(3) + 1
                                totEff_Inc(3) = totEff_Inc(3) + Cells(Data_i, 13).Value
                                cloDur_Inc(3) = cloDur_Inc(3) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Inc(3) = resSla_Inc(3) + 1
                                End If
                            End If
                        End If
'************* Case 4 and Case 5 both sharing same variables ***************
                    Case 5
                        If createDate < startDate Then
                            opnBal_Inc(3) = opnBal_Inc(3) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Inc(3) = carried_Inc(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Inc(3) = closed_Inc(3) + 1
                                totEff_Inc(3) = totEff_Inc(3) + Cells(Data_i, 13).Value
                                cloDur_Inc(3) = cloDur_Inc(3) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Inc(3) = resSla_Inc(3) + 1
                                End If
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Inc(3) = recvd_Inc(3) + 1
                            If rspSLA = "N" Then
                                rspSla_Inc(3) = rspSla_Inc(3) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Inc(3) = reOpen_Inc(3) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Inc(3) = carried_Inc(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Inc(3) = closed_Inc(3) + 1
                                totEff_Inc(3) = totEff_Inc(3) + Cells(Data_i, 13).Value
                                cloDur_Inc(3) = cloDur_Inc(3) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Inc(3) = resSla_Inc(3) + 1
                                End If
                            End If
                        End If
                End Select
'If ticket type = SRQ
                Case "SRQ"
                    If SRQ_Dict.Exists(element) Then
                        SRQ_Dict.Item(element) = SRQ_Dict.Item(element) + 1
                    Else
                        SRQ_Dict.Add element, 1
                    End If
                Select Case prty
                    'Priority 1 and so on same for below also
                    Case 1
                        '------------------ Filtering Data for Quarter ----------------------
                        'Opening balance: finish date = "" and create Date <= Start Date
                        If createDate < startDate Then
                            'Carried Forward & Opening Balance:
                            opnBal_Srq(0) = opnBal_Srq(0) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Srq(0) = carried_Srq(0) + 1
                            'Resolved: end_Date >= 'Finish Date' >= start_Date
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Srq(0) = closed_Srq(0) + 1
                                'Total Effort Spent
                                totEff_Srq(0) = totEff_Srq(0) + Cells(Data_i, 13).Value
                                'Total Closure Duration
                                cloDur_Srq(0) = cloDur_Srq(0) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    'Total Resolution SLA Breached
                                    resSla_Srq(0) = resSla_Srq(0) + 1
                                End If
                            End If
                        'Received: end_Date >= 'Create Date' >= start_Date
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Srq(0) = recvd_Srq(0) + 1
                            If rspSLA = "N" Then
                                'Total Response SLA Breached
                                rspSla_Srq(0) = rspSla_Srq(0) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Srq(0) = reOpen_Srq(0) + 1
                            End If
                            'Carried Forward:
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Srq(0) = carried_Srq(0) + 1
                            'Resolved: end_Date >= 'Finish Date' >= start_Date
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Srq(0) = closed_Srq(0) + 1
                                'Total Effort Spent
                                totEff_Srq(0) = totEff_Srq(0) + Cells(Data_i, 13).Value
                                'Total Closure Duration
                                cloDur_Srq(0) = cloDur_Srq(0) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    'Total Resolution SLA Breached
                                    resSla_Srq(0) = resSla_Srq(0) + 1
                                End If
                            End If
                        End If
                    Case 2
                        If createDate < startDate Then
                            opnBal_Srq(1) = opnBal_Srq(1) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Srq(1) = carried_Srq(1) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Srq(1) = closed_Srq(1) + 1
                                totEff_Srq(1) = totEff_Srq(1) + Cells(Data_i, 13).Value
                                cloDur_Srq(1) = cloDur_Srq(1) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Srq(1) = resSla_Srq(1) + 1
                                End If
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Srq(1) = recvd_Srq(1) + 1
                            If rspSLA = "N" Then
                                rspSla_Srq(1) = rspSla_Srq(1) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Srq(1) = reOpen_Srq(1) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Srq(1) = carried_Srq(1) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Srq(1) = closed_Srq(1) + 1
                                totEff_Srq(1) = totEff_Srq(1) + Cells(Data_i, 13).Value
                                cloDur_Srq(1) = cloDur_Srq(1) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Srq(1) = resSla_Srq(1) + 1
                                End If
                            End If
                        End If
                    Case 3
                        If createDate < startDate Then
                            opnBal_Srq(2) = opnBal_Srq(2) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Srq(2) = carried_Srq(2) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Srq(2) = closed_Srq(2) + 1
                                totEff_Srq(2) = totEff_Srq(2) + Cells(Data_i, 13).Value
                                cloDur_Srq(2) = cloDur_Srq(2) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Srq(2) = resSla_Srq(2) + 1
                                End If
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Srq(2) = recvd_Srq(2) + 1
                            If rspSLA = "N" Then
                                rspSla_Srq(2) = rspSla_Srq(2) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Srq(2) = reOpen_Srq(2) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Srq(2) = carried_Srq(2) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Srq(2) = closed_Srq(2) + 1
                                totEff_Srq(2) = totEff_Srq(2) + Cells(Data_i, 13).Value
                                cloDur_Srq(2) = cloDur_Srq(2) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Srq(2) = resSla_Srq(2) + 1
                                End If
                            End If
                        End If
                    Case 4
                        If createDate < startDate Then
                            opnBal_Srq(3) = opnBal_Srq(3) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Srq(3) = carried_Srq(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Srq(3) = closed_Srq(3) + 1
                                totEff_Srq(3) = totEff_Srq(3) + Cells(Data_i, 13).Value
                                cloDur_Srq(3) = cloDur_Srq(3) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Srq(3) = resSla_Srq(3) + 1
                                End If
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Srq(3) = recvd_Srq(3) + 1
                            If rspSLA = "N" Then
                                rspSla_Srq(3) = rspSla_Srq(3) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Srq(3) = reOpen_Srq(3) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Srq(3) = carried_Srq(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Srq(3) = closed_Srq(3) + 1
                                totEff_Srq(3) = totEff_Srq(3) + Cells(Data_i, 13).Value
                                cloDur_Srq(3) = cloDur_Srq(3) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Srq(3) = resSla_Srq(3) + 1
                                End If
                            End If
                        End If
'************* Case 4 and Case 5 both sharing same variables ***************
                    Case 5
                        If createDate < startDate Then
                            opnBal_Srq(3) = opnBal_Srq(3) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Srq(3) = carried_Srq(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Srq(3) = closed_Srq(3) + 1
                                totEff_Srq(3) = totEff_Srq(3) + Cells(Data_i, 13).Value
                                cloDur_Srq(3) = cloDur_Srq(3) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Srq(3) = resSla_Srq(3) + 1
                                End If
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Srq(3) = recvd_Srq(3) + 1
                            If rspSLA = "N" Then
                                rspSla_Srq(3) = rspSla_Srq(3) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Srq(3) = reOpen_Srq(3) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Srq(3) = carried_Srq(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Srq(3) = closed_Srq(3) + 1
                                totEff_Srq(3) = totEff_Srq(3) + Cells(Data_i, 13).Value
                                cloDur_Srq(3) = cloDur_Srq(3) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Srq(3) = resSla_Srq(3) + 1
                                End If
                            End If
                        End If
                    End Select
'If ticket type = PRB
                Case "PRB"
                    If PRB_Dict.Exists(element) Then
                        PRB_Dict.Item(element) = PRB_Dict.Item(element) + 1
                    Else
                        PRB_Dict.Add element, 1
                    End If
                Select Case prty
                    'Priority 1 and so on same for below also
                    Case 1
                        '------------------ Filtering Data for Quarter ----------------------
                        'Opening balance: finish date = "" and create Date <= Start Date
                        If createDate < startDate Then
                            'Carried Forward & Opening Balance:
                            opnBal_Prb(0) = opnBal_Prb(0) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Prb(0) = carried_Prb(0) + 1
                            'Resolved: end_Date >= 'Finish Date' >= start_Date
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Prb(0) = closed_Prb(0) + 1
                                'Total Effort Spent
                                totEff_Prb(0) = totEff_Prb(0) + Cells(Data_i, 13).Value
                                'Total Closure Duration
                                cloDur_Prb(0) = cloDur_Prb(0) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    'Total Resolution SLA Breached
                                    resSla_Prb(0) = resSla_Prb(0) + 1
                                End If
                            End If
                        'Received: end_Date >= 'Create Date' >= start_Date
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Prb(0) = recvd_Prb(0) + 1
                            If rspSLA = "N" Then
                                'Total Response SLA Breached
                                rspSla_Prb(0) = rspSla_Prb(0) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Prb(0) = reOpen_Prb(0) + 1
                            End If
                            'Carried Forward:
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Prb(0) = carried_Prb(0) + 1
                            'Resolved: end_Date >= 'Finish Date' >= start_Date
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Prb(0) = closed_Prb(0) + 1
                                'Total Effort Spent
                                totEff_Prb(0) = totEff_Prb(0) + Cells(Data_i, 13).Value
                                'Total Closure Duration
                                cloDur_Prb(0) = cloDur_Prb(0) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    'Total Resolution SLA Breached
                                    resSla_Prb(0) = resSla_Prb(0) + 1
                                End If
                            End If
                        End If
                    Case 2
                        If createDate < startDate Then
                            opnBal_Prb(1) = opnBal_Prb(1) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Prb(1) = carried_Prb(1) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Prb(1) = closed_Prb(1) + 1
                                totEff_Prb(1) = totEff_Prb(1) + Cells(Data_i, 13).Value
                                cloDur_Prb(1) = cloDur_Prb(1) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Prb(1) = resSla_Prb(1) + 1
                                End If
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Prb(1) = recvd_Prb(1) + 1
                            If rspSLA = "N" Then
                                rspSla_Prb(1) = rspSla_Prb(1) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Prb(1) = reOpen_Prb(1) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Prb(1) = carried_Prb(1) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Prb(1) = closed_Prb(1) + 1
                                totEff_Prb(1) = totEff_Prb(1) + Cells(Data_i, 13).Value
                                cloDur_Prb(1) = cloDur_Prb(1) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Prb(1) = resSla_Prb(1) + 1
                                End If
                            End If
                        End If
                    Case 3
                        If createDate < startDate Then
                            opnBal_Prb(2) = opnBal_Prb(2) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Prb(2) = carried_Prb(2) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Prb(2) = closed_Prb(2) + 1
                                totEff_Prb(2) = totEff_Prb(2) + Cells(Data_i, 13).Value
                                cloDur_Prb(2) = cloDur_Prb(2) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Prb(2) = resSla_Prb(2) + 1
                                End If
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Prb(2) = recvd_Prb(2) + 1
                            If rspSLA = "N" Then
                                rspSla_Prb(2) = rspSla_Prb(2) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Prb(2) = reOpen_Prb(2) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Prb(2) = carried_Prb(2) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Prb(2) = closed_Prb(2) + 1
                                totEff_Prb(2) = totEff_Prb(2) + Cells(Data_i, 13).Value
                                cloDur_Prb(2) = cloDur_Prb(2) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Prb(2) = resSla_Prb(2) + 1
                                End If
                            End If
                        End If
                    Case 4
                        If createDate < startDate Then
                            opnBal_Prb(3) = opnBal_Prb(3) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Prb(3) = carried_Prb(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Prb(3) = closed_Prb(3) + 1
                                totEff_Prb(3) = totEff_Prb(3) + Cells(Data_i, 13).Value
                                cloDur_Prb(3) = cloDur_Prb(3) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Prb(3) = resSla_Prb(3) + 1
                                End If
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Prb(3) = recvd_Prb(3) + 1
                            If rspSLA = "N" Then
                                rspSla_Prb(3) = rspSla_Prb(3) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Prb(3) = reOpen_Prb(3) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Prb(3) = carried_Prb(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Prb(3) = closed_Prb(3) + 1
                                totEff_Prb(3) = totEff_Prb(3) + Cells(Data_i, 13).Value
                                cloDur_Prb(3) = cloDur_Prb(3) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Prb(3) = resSla_Prb(3) + 1
                                End If
                            End If
                        End If
'************* Case 4 and Case 5 both sharing same variables ***************
                    Case 5
                        If createDate < startDate Then
                            opnBal_Prb(3) = opnBal_Prb(3) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Prb(3) = carried_Prb(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Prb(3) = closed_Prb(3) + 1
                                totEff_Prb(3) = totEff_Prb(3) + Cells(Data_i, 13).Value
                                cloDur_Prb(3) = cloDur_Prb(3) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Prb(3) = resSla_Prb(3) + 1
                                End If
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Prb(3) = recvd_Prb(3) + 1
                            If rspSLA = "N" Then
                                rspSla_Prb(3) = rspSla_Prb(3) + 1
                            End If
                            If reOpnd = "Y" Then
                                reOpen_Prb(3) = reOpen_Prb(3) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Prb(3) = carried_Prb(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Prb(3) = closed_Prb(3) + 1
                                totEff_Prb(3) = totEff_Prb(3) + Cells(Data_i, 13).Value
                                cloDur_Prb(3) = cloDur_Prb(3) + Cells(Data_i, 19).Value
                                If resSLA = "N" Then
                                    resSla_Prb(3) = resSla_Prb(3) + 1
                                End If
                            End If
                        End If
                End Select
                'If ticket type = CHG
                Case "ACT"
                    If CHG_Dict.Exists(element) Then
                        CHG_Dict.Item(element) = CHG_Dict.Item(element) + 1
                    Else
                        CHG_Dict.Add element, 1
                    End If
                Select Case prty
                    'Priority 1 and so on same for below also
                    Case 1
                        '------------------ Filtering Data for Quarter ----------------------
                        'Opening balance: finish date = "" and create Date <= Start Date
                        If createDate < startDate Then
                            'Carried Forward & Opening Balance:
                            opnBal_Chg(0) = opnBal_Chg(0) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Chg(0) = carried_Chg(0) + 1
                            'Resolved: end_Date >= 'Finish Date' >= start_Date
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Chg(0) = closed_Chg(0) + 1
                                'Total Effort Spent
                                totEff_Chg(0) = totEff_Chg(0) + Cells(Data_i, 13).Value
                            End If
                        'Received: end_Date >= 'Create Date' >= start_Date
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Chg(0) = recvd_Chg(0) + 1
                            If reOpnd = "Y" Then
                                reOpen_Chg(0) = reOpen_Chg(0) + 1
                            End If
                            'Carried Forward:
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Chg(0) = carried_Chg(0) + 1
                            'Resolved: end_Date >= 'Finish Date' >= start_Date
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Chg(0) = closed_Chg(0) + 1
                                'Total Effort Spent
                                totEff_Chg(0) = totEff_Chg(0) + Cells(Data_i, 13).Value
                                'Change Window missed implemented
                                If finishDate < Cells(Data_i, 16).Value And finishDate > Cells(Data_i, 17).Value Then
                                    winMiss_Chg(0) = winMiss_Chg(0) + 1
                                End If
                            End If
                        End If
                    Case 2
                        If createDate < startDate Then
                            opnBal_Chg(1) = opnBal_Chg(1) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Chg(1) = carried_Chg(1) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Chg(1) = closed_Chg(1) + 1
                                totEff_Chg(1) = totEff_Chg(1) + Cells(Data_i, 13).Value
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Chg(1) = recvd_Chg(1) + 1
                            If reOpnd = "Y" Then
                                reOpen_Chg(1) = reOpen_Chg(1) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Chg(1) = carried_Chg(1) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Chg(1) = closed_Chg(1) + 1
                                totEff_Chg(1) = totEff_Chg(1) + Cells(Data_i, 13).Value
                                If finishDate < Cells(Data_i, 16).Value And finishDate > Cells(Data_i, 17).Value Then
                                    winMiss_Chg(1) = winMiss_Chg(1) + 1
                                End If
                            End If
                        End If
                    Case 3
                        If createDate < startDate Then
                            opnBal_Chg(2) = opnBal_Chg(2) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Chg(2) = carried_Chg(2) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Chg(2) = closed_Chg(2) + 1
                                totEff_Chg(2) = totEff_Chg(2) + Cells(Data_i, 13).Value
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Chg(2) = recvd_Chg(2) + 1
                            If reOpnd = "Y" Then
                                reOpen_Chg(2) = reOpen_Chg(2) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Chg(2) = carried_Chg(2) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Chg(2) = closed_Chg(2) + 1
                                totEff_Chg(2) = totEff_Chg(2) + Cells(Data_i, 13).Value
                                If finishDate < Cells(Data_i, 16).Value And finishDate > Cells(Data_i, 17).Value Then
                                    winMiss_Chg(2) = winMiss_Chg(2) + 1
                                End If
                            End If
                        End If
                    Case 4
                        If createDate < startDate Then
                            opnBal_Chg(3) = opnBal_Chg(3) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Chg(3) = carried_Chg(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Chg(3) = closed_Chg(3) + 1
                                totEff_Chg(3) = totEff_Chg(3) + Cells(Data_i, 13).Value
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Chg(3) = recvd_Chg(3) + 1
                            If reOpnd = "Y" Then
                                reOpen_Chg(3) = reOpen_Chg(3) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Chg(3) = carried_Chg(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Chg(3) = closed_Chg(3) + 1
                                totEff_Chg(3) = totEff_Chg(3) + Cells(Data_i, 13).Value
                                If finishDate < Cells(Data_i, 16).Value And finishDate > Cells(Data_i, 17).Value Then
                                    winMiss_Chg(3) = winMiss_Chg(0) + 1
                                End If
                            End If
                        End If
'************* Case 4 and Case 5 both sharing same variables ***************
                    Case 5
                        If createDate < startDate Then
                            opnBal_Chg(3) = opnBal_Chg(3) + 1
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Chg(3) = carried_Chg(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Chg(3) = closed_Chg(3) + 1
                                totEff_Chg(3) = totEff_Chg(3) + Cells(Data_i, 13).Value
                            End If
                        ElseIf createDate >= startDate And finishDate <= endDate Then
                            recvd_Chg(3) = recvd_Chg(3) + 1
                            If reOpnd = "Y" Then
                                reOpen_Chg(3) = reOpen_Chg(3) + 1
                            End If
                            If CStr(finishDate) = "" Or finishDate > endDate Then
                                carried_Chg(3) = carried_Chg(3) + 1
                            ElseIf finishDate >= startDate And finishDate <= endDate Then
                                closed_Chg(3) = closed_Chg(3) + 1
                                totEff_Chg(3) = totEff_Chg(3) + Cells(Data_i, 13).Value
                                If finishDate < Cells(Data_i, 16).Value And finishDate > Cells(Data_i, 17).Value Then
                                    winMiss_Chg(3) = winMiss_Chg(3) + 1
                                End If
                            End If
                        End If
                End Select
            End Select
    Next Data_i
    ' Total of the following variables are calculated
    For i = 0 To 3
        opnBal_Inc(4) = opnBal_Inc(4) + opnBal_Inc(i)
        recvd_Inc(4) = recvd_Inc(4) + recvd_Inc(i)
        carried_Inc(4) = carried_Inc(4) + carried_Inc(i)
        closed_Inc(4) = closed_Inc(4) + closed_Inc(i)
        reOpen_Inc(4) = reOpen_Inc(4) + reOpen_Inc(i)
        totEff_Inc(4) = totEff_Inc(4) + totEff_Inc(i)
        rspSla_Inc(4) = rspSla_Inc(4) + rspSla_Inc(i)
        resSla_Inc(4) = resSla_Inc(4) + resSla_Inc(i)
        cloDur_Inc(4) = cloDur_Inc(4) + cloDur_Inc(i)
        
        opnBal_Srq(4) = opnBal_Srq(4) + opnBal_Srq(i)
        recvd_Srq(4) = recvd_Srq(4) + recvd_Srq(i)
        carried_Srq(4) = carried_Srq(4) + carried_Srq(i)
        closed_Srq(4) = closed_Srq(4) + closed_Srq(i)
        reOpen_Srq(4) = reOpen_Srq(4) + reOpen_Srq(i)
        totEff_Srq(4) = totEff_Srq(4) + totEff_Srq(i)
        rspSla_Srq(4) = rspSla_Srq(4) + rspSla_Srq(i)
        resSla_Srq(4) = resSla_Srq(4) + resSla_Srq(i)
        cloDur_Srq(4) = cloDur_Srq(4) + cloDur_Srq(i)
        
        opnBal_Prb(4) = opnBal_Prb(4) + opnBal_Prb(i)
        recvd_Prb(4) = recvd_Prb(4) + recvd_Prb(i)
        carried_Prb(4) = carried_Prb(4) + carried_Prb(i)
        closed_Prb(4) = closed_Prb(4) + closed_Prb(i)
        reOpen_Prb(4) = reOpen_Prb(4) + reOpen_Prb(i)
        totEff_Prb(4) = totEff_Prb(4) + totEff_Prb(i)
        rspSla_Prb(4) = rspSla_Prb(4) + rspSla_Prb(i)
        resSla_Prb(4) = resSla_Prb(4) + resSla_Prb(i)
        cloDur_Prb(4) = cloDur_Prb(4) + cloDur_Prb(i)
        
        opnBal_Chg(4) = opnBal_Chg(4) + opnBal_Chg(i)
        recvd_Chg(4) = recvd_Chg(4) + recvd_Chg(i)
        carried_Chg(4) = carried_Chg(4) + carried_Chg(i)
        closed_Chg(4) = closed_Chg(4) + closed_Chg(i)
        reOpen_Chg(4) = reOpen_Chg(4) + reOpen_Chg(i)
        totEff_Chg(4) = totEff_Chg(4) + totEff_Chg(i)
        winMiss_Chg(4) = winMiss_Chg(4) + winMiss_Chg(i)
    Next i
    'Conversion of Minutes to Hours for Total Effort
    For i = 0 To 4
        totEff_Inc(i) = Round(totEff_Inc(i) / 60, 2)
        totEff_Srq(i) = Round(totEff_Srq(i) / 60, 2)
        totEff_Prb(i) = Round(totEff_Prb(i) / 60, 2)
        totEff_Chg(i) = Round(totEff_Chg(i) / 60, 2)
    Next i
    'average is being calculated here
    For i = 0 To 4
        If closed_Inc(i) <> 0 Then
            avgEff_Inc(i) = Round(totEff_Inc(i) / closed_Inc(i), 2)
            resSlaPrcnt_Inc(i) = resSla_Inc(i) * 100 / closed_Inc(i)
            avgClosure_Inc(i) = cloDur_Inc(i) / closed_Inc(i)
        End If
        If closed_Srq(i) <> 0 Then
            avgEff_Srq(i) = Round(totEff_Srq(i) / closed_Srq(i), 2)
            resSlaPrcnt_Srq(i) = resSla_Srq(i) * 100 / closed_Srq(i)
            avgClosure_Srq(i) = cloDur_Srq(i) / closed_Srq(i)
        End If
        If closed_Prb(i) <> 0 Then
            avgEff_Prb(i) = Round(totEff_Prb(i) / closed_Prb(i), 2)
            resSlaPrcnt_Prb(i) = resSla_Prb(i) * 100 / closed_Prb(i)
            avgClosure_Prb(i) = cloDur_Prb(i) / closed_Prb(i)
        End If
        If closed_Chg(i) <> 0 Then
            avgEff_Chg(i) = Round(totEff_Chg(i) / closed_Chg(i), 2)
            winMissPrcnt_Chg(i) = winMiss_Chg(i) * 100 / closed_Chg(i)
        End If
        
        If recvd_Inc(i) <> 0 Then
            rspSlaPrcnt_Inc(i) = rspSla_Inc(i) * 100 / recvd_Inc(i)
        End If
        If recvd_Srq(i) <> 0 Then
            rspSlaPrcnt_Srq(i) = rspSla_Srq(i) * 100 / recvd_Srq(i)
        End If
        If recvd_Prb(i) <> 0 Then
            rspSlaPrcnt_Prb(i) = rspSla_Prb(i) * 100 / recvd_Prb(i)
        End If
    Next i
    
    'Checking Whether Dictionary contains Blank String
    If INC_Dict.Exists("") Then
        tmSize_Inc = INC_Dict.Count - 1
    End If
    If SRQ_Dict.Exists("") Then
        tmSize_Srq = SRQ_Dict.Count - 1
    End If
    If CHG_Dict.Exists("") Then
        tmSize_Chg = CHG_Dict.Count - 1
    End If
    If PRB_Dict.Exists("") Then
        tmSize_Prb = PRB_Dict.Count - 1
    End If
    
    
    WS_CSS.Select
    
    '------------------ VERSIONWISE Value Placement of the Variable in Excel sheet -------------
    '----------Incident----------
    Range("D" & 34 + 15 * v & ":H" & 34 + 15 * v).Value = opnBal_Inc
    Range("D" & 35 + 15 * v & ":H" & 35 + 15 * v).Value = recvd_Inc
    Range("D" & 36 + 15 * v & ":H" & 36 + 15 * v).Value = carried_Inc
    Range("D" & 37 + 15 * v & ":H" & 37 + 15 * v).Value = closed_Inc
    Range("D" & 38 + 15 * v & ":H" & 38 + 15 * v).Value = reOpen_Inc
    Range("D" & 39 + 15 * v & ":H" & 39 + 15 * v).Value = totEff_Inc
    Range("D" & 40 + 15 * v & ":H" & 40 + 15 * v).Value = avgEff_Inc
    Cells(41 + 15 * v, 4).Value = tmSize_Inc
    Range("D" & 44 + 15 * v & ":H" & 44 + 15 * v).Value = rspSla_Inc
    Range("D" & 45 + 15 * v & ":H" & 45 + 15 * v).Value = rspSlaPrcnt_Inc
    Range("D" & 46 + 15 * v & ":H" & 46 + 15 * v).Value = resSla_Inc
    Range("D" & 47 + 15 * v & ":H" & 47 + 15 * v).Value = resSlaPrcnt_Inc
    Range("D" & 48 + 15 * v & ":H" & 48 + 15 * v).Value = avgClosure_Inc
    '----------Service Request----------
    Range("I" & 34 + 15 * v & ":M" & 34 + 15 * v).Value = opnBal_Srq
    Range("I" & 35 + 15 * v & ":M" & 35 + 15 * v).Value = recvd_Srq
    Range("I" & 36 + 15 * v & ":M" & 36 + 15 * v).Value = carried_Srq
    Range("I" & 37 + 15 * v & ":M" & 37 + 15 * v).Value = closed_Srq
    Range("I" & 38 + 15 * v & ":M" & 38 + 15 * v).Value = reOpen_Srq
    Range("I" & 39 + 15 * v & ":M" & 39 + 15 * v).Value = totEff_Srq
    Range("I" & 40 + 15 * v & ":M" & 40 + 15 * v).Value = avgEff_Srq
    Cells(41 + 15 * v, 9).Value = tmSize_Srq
    Range("I" & 44 + 15 * v & ":M" & 44 + 15 * v).Value = rspSla_Srq
    Range("I" & 45 + 15 * v & ":M" & 45 + 15 * v).Value = rspSlaPrcnt_Srq
    Range("I" & 46 + 15 * v & ":M" & 46 + 15 * v).Value = resSla_Srq
    Range("I" & 47 + 15 * v & ":M" & 47 + 15 * v).Value = resSlaPrcnt_Srq
    Range("I" & 48 + 15 * v & ":M" & 48 + 15 * v).Value = avgClosure_Srq
    '----------Problem----------
    Range("N" & 34 + 15 * v & ":R" & 34 + 15 * v).Value = opnBal_Prb
    Range("N" & 35 + 15 * v & ":R" & 35 + 15 * v).Value = recvd_Prb
    Range("N" & 36 + 15 * v & ":R" & 36 + 15 * v).Value = carried_Prb
    Range("N" & 37 + 15 * v & ":R" & 37 + 15 * v).Value = closed_Prb
    Range("N" & 38 + 15 * v & ":R" & 38 + 15 * v).Value = reOpen_Prb
    Range("N" & 39 + 15 * v & ":R" & 39 + 15 * v).Value = totEff_Prb
    Range("N" & 40 + 15 * v & ":R" & 40 + 15 * v).Value = avgEff_Prb
    Cells(41 + 15 * v, 14).Value = tmSize_Prb
    Range("N" & 44 + 15 * v & ":R" & 44 + 15 * v).Value = rspSla_Prb
    Range("N" & 45 + 15 * v & ":R" & 45 + 15 * v).Value = rspSlaPrcnt_Prb
    Range("N" & 46 + 15 * v & ":R" & 46 + 15 * v).Value = resSla_Prb
    Range("N" & 47 + 15 * v & ":R" & 47 + 15 * v).Value = resSlaPrcnt_Prb
    Range("N" & 48 + 15 * v & ":R" & 48 + 15 * v).Value = avgClosure_Prb
    '----------Change Req----------
    Range("S" & 34 + 15 * v & ":W" & 34 + 15 * v).Value = opnBal_Chg
    Range("S" & 35 + 15 * v & ":W" & 35 + 15 * v).Value = recvd_Chg
    Range("S" & 36 + 15 * v & ":W" & 36 + 15 * v).Value = carried_Chg
    Range("S" & 37 + 15 * v & ":W" & 37 + 15 * v).Value = closed_Chg
    Range("S" & 38 + 15 * v & ":W" & 38 + 15 * v).Value = reOpen_Chg
    Range("S" & 39 + 15 * v & ":W" & 39 + 15 * v).Value = totEff_Chg
    Range("S" & 40 + 15 * v & ":W" & 40 + 15 * v).Value = avgEff_Chg
    Cells(41 + 15 * v, 19).Value = tmSize_Chg
    Range("S" & 42 + 15 * v & ":W" & 42 + 15 * v).Value = winMiss_Chg
    Range("S" & 43 + 15 * v & ":W" & 43 + 15 * v).Value = winMissPrcnt_Chg
    'Quarter Range of date placing
    Cells(34 + 15 * v, 2).Value = quarters(v, 0) & "  -  " & quarters(v, 1)
    
End Sub
