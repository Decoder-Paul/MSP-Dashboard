Attribute VB_Name = "Aging"
Sub agingCount(ByVal team As String)
'========================================================================================================
' Main Data for Staging
' -------------------------------------------------------------------------------------------------------
' Purpose   :   Calculating aging for the Active tickets
'
' Author    :   Shambhavi B M, 11th January, 2018
' Notes     :   Different Ticket Types: 'INC', 'SRQ', 'PRB' are string constant
'
' Parameter :   N/A
' Returns   :   N/A
' -------------------------------------------------------------------------------------------------------
' Revision History
'
'========================================================================================================

'----------Incident----------
Dim Days0_1_INC(4) As Integer
Dim Days2_3_INC(4) As Integer
Dim Days4_5_INC(4) As Integer
Dim Days6_7_INC(4) As Integer
Dim Days8_14_INC(4) As Integer
Dim Days15_30_INC(4) As Integer
Dim Days31_60_INC(4) As Integer
Dim Days61_90_INC(4) As Integer
Dim Days_GT_90_INC(4) As Integer
Dim Days_Sum_INC(4) As Integer

'--------Service Request--------
Dim Days0_1_SRQ(4) As Integer
Dim Days2_3_SRQ(4) As Integer
Dim Days4_5_SRQ(4) As Integer
Dim Days6_7_SRQ(4) As Integer
Dim Days8_14_SRQ(4) As Integer
Dim Days15_30_SRQ(4) As Integer
Dim Days31_60_SRQ(4) As Integer
Dim Days61_90_SRQ(4) As Integer
Dim Days_GT_90_SRQ(4) As Integer
Dim Days_Sum_SRQ(4) As Integer

'--------Problem Statement--------
Dim Days0_1_PRB(4) As Integer
Dim Days2_3_PRB(4) As Integer
Dim Days4_5_PRB(4) As Integer
Dim Days6_7_PRB(4) As Integer
Dim Days8_14_PRB(4) As Integer
Dim Days15_30_PRB(4) As Integer
Dim Days31_60_PRB(4) As Integer
Dim Days61_90_PRB(4) As Integer
Dim Days_GT_90_PRB(4) As Integer
Dim Days_Sum_PRB(4) As Integer

Dim today As Variant
Dim DAlro As Long
Dim i As Long
Dim ticket_type As String
Dim priority As Long
Dim age_of_tkt As Variant

WS_CSS.Activate
today = DateOfreport

WS_DA.Activate
DAlro = WS_DA.Cells(WS_DA.Rows.Count, "A").End(xlUp).Row

'cleaning Aging column in mainData before calculating aging for team
WS_DA.Range("S2:S" & DAlro).Clear

'Aging calculation logic
For i = 2 To DAlro

    'checking whether Actualfinish date is not empty and should be greater than date of report
    If Cells(i, 25).Value <> "" Or Cells(i, 25).Value >= CLng(DateOfreport) Then
        'checking whether Actual start date is not empty or not
        If Cells(i, 24).Value <> "" Then
                'if Actual start date is empty then taking dufference between DateOfreport and Actualstartdate
                Cells(i, 19).Value = CLng(DateOfreport) - Cells(i, 24).Value
        Else
            'if Actual start date is empty then taking the difference between DateOfreport and Creation date
            If Cells(i, 24).Value = "" Then
                    Cells(i, 19).Value = CLng(DateOfreport) - Cells(i, 23).Value
            End If
        End If
    Else
        'if Actualfinish date empty aging is empty
        Cells(i, 19).Value = ""
    End If
Next i

'------------Counting aging based on team, ticket type and priority-----------------
For i = 2 To DAlro
    
     If Cells(i, 8).Value = team Then
             
        ticket_type = Cells(i, 1).Value
        priority = Cells(i, 12).Value
        age_of_tkt = Cells(i, 19).Value

        Select Case ticket_type
            Case "INC":
                Select Case priority
                    Case 1:
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_INC(0) = Days0_1_INC(0) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_INC(0) = Days2_3_INC(0) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_INC(0) = Days4_5_INC(0) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_INC(0) = Days6_7_INC(0) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_INC(0) = Days8_14_INC(0) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_INC(0) = Days15_30_INC(0) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_INC(0) = Days31_60_INC(0) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_INC(0) = Days61_90_INC(0) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_INC(0) = Days_GT_90_INC(0) + 1
                                End If
                            End If
                    Case 2:
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_INC(1) = Days0_1_INC(1) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_INC(1) = Days2_3_INC(1) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_INC(1) = Days4_5_INC(1) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_INC(1) = Days6_7_INC(1) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_INC(1) = Days8_14_INC(1) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_INC(1) = Days15_30_INC(1) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_INC(1) = Days31_60_INC(1) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_INC(1) = Days61_90_INC(1) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_INC(1) = Days_GT_90_INC(1) + 1
                                End If
                            End If
                            
                    Case 3:
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_INC(2) = Days0_1_INC(2) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_INC(2) = Days2_3_INC(2) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_INC(2) = Days4_5_INC(2) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_INC(2) = Days6_7_INC(2) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_INC(2) = Days8_14_INC(2) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_INC(2) = Days15_30_INC(2) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_INC(2) = Days31_60_INC(2) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_INC(2) = Days61_90_INC(2) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_INC(2) = Days_GT_90_INC(2) + 1
                                End If
                            End If
                    
                    Case 4 And 5:
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_INC(3) = Days0_1_INC(3) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_INC(3) = Days2_3_INC(3) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_INC(3) = Days4_5_INC(3) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_INC(3) = Days6_7_INC(3) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_INC(3) = Days8_14_INC(3) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_INC(3) = Days15_30_INC(3) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_INC(3) = Days31_60_INC(3) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_INC(3) = Days61_90_INC(3) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_INC(3) = Days_GT_90_INC(3) + 1
                                End If
                            End If
                                
                End Select
            Case "SRQ":
                Select Case priority
                    Case 1:
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_SRQ(0) = Days0_1_SRQ(0) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_SRQ(0) = Days2_3_SRQ(0) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_SRQ(0) = Days4_5_SRQ(0) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_SRQ(0) = Days6_7_SRQ(0) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_SRQ(0) = Days8_14_SRQ(0) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_SRQ(0) = Days15_30_SRQ(0) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_SRQ(0) = Days31_60_SRQ(0) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_SRQ(0) = Days61_90_SRQ(0) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_SRQ(0) = Days_GT_90_SRQ(0) + 1
                                End If
                            End If
                    Case 2:
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_SRQ(1) = Days0_1_SRQ(1) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_SRQ(1) = Days2_3_SRQ(1) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_SRQ(1) = Days4_5_SRQ(1) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_SRQ(1) = Days6_7_SRQ(1) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_SRQ(1) = Days8_14_SRQ(1) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_SRQ(1) = Days15_30_SRQ(1) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_SRQ(1) = Days31_60_SRQ(1) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_SRQ(1) = Days61_90_SRQ(1) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_SRQ(1) = Days_GT_90_SRQ(1) + 1
                                End If
                            End If
                            
                    Case 3:
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_SRQ(2) = Days0_1_SRQ(2) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_SRQ(2) = Days2_3_SRQ(2) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_SRQ(2) = Days4_5_SRQ(2) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_SRQ(2) = Days6_7_SRQ(2) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_SRQ(2) = Days8_14_SRQ(2) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_SRQ(2) = Days15_30_SRQ(2) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_SRQ(2) = Days31_60_SRQ(2) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_SRQ(2) = Days61_90_SRQ(2) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_SRQ(2) = Days_GT_90_SRQ(2) + 1
                                End If
                            End If
                    
                    Case 4 And 5:
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_SRQ(3) = Days0_1_SRQ(3) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_SRQ(3) = Days2_3_SRQ(3) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_SRQ(3) = Days4_5_SRQ(3) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_SRQ(3) = Days6_7_SRQ(3) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_SRQ(3) = Days8_14_SRQ(3) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_SRQ(3) = Days15_30_SRQ(3) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_SRQ(3) = Days31_60_SRQ(3) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_SRQ(3) = Days61_90_SRQ(3) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_SRQ(3) = Days_GT_90_SRQ(3) + 1
                                End If
                            End If
                End Select
            Case "PRB":
                Select Case priority
                    Case 1:
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_PRB(0) = Days0_1_PRB(0) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_PRB(0) = Days2_3_PRB(0) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_PRB(0) = Days4_5_PRB(0) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_PRB(0) = Days6_7_PRB(0) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_PRB(0) = Days8_14_PRB(0) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_PRB(0) = Days15_30_PRB(0) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_PRB(0) = Days31_60_PRB(0) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_PRB(0) = Days61_90_PRB(0) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_PRB(0) = Days_GT_90_PRB(0) + 1
                                End If
                            End If
                    Case 2:
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_PRB(1) = Days0_1_PRB(1) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_PRB(1) = Days2_3_PRB(1) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_PRB(1) = Days4_5_PRB(1) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_PRB(1) = Days6_7_PRB(1) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_PRB(1) = Days8_14_PRB(1) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_PRB(1) = Days15_30_PRB(1) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_PRB(1) = Days31_60_PRB(1) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_PRB(1) = Days61_90_PRB(1) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_PRB(1) = Days_GT_90_PRB(1) + 1
                                End If
                            End If
                            
                    Case 3:
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_PRB(2) = Days0_1_PRB(2) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_PRB(2) = Days2_3_PRB(2) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_PRB(2) = Days4_5_PRB(2) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_PRB(2) = Days6_7_PRB(2) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_PRB(2) = Days8_14_PRB(2) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_PRB(2) = Days15_30_PRB(2) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_PRB(2) = Days31_60_PRB(2) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_PRB(2) = Days61_90_PRB(2) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_PRB(2) = Days_GT_90_PRB(2) + 1
                                End If
                            End If
                    
                    Case 4 And 5:
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_PRB(3) = Days0_1_PRB(3) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_PRB(3) = Days2_3_PRB(3) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_PRB(3) = Days4_5_PRB(3) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_PRB(3) = Days6_7_PRB(3) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_PRB(3) = Days8_14_PRB(3) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_PRB(3) = Days15_30_PRB(3) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_PRB(3) = Days31_60_PRB(3) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_PRB(3) = Days61_90_PRB(3) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_PRB(3) = Days_GT_90_PRB(3) + 1
                                End If
                            End If
                End Select
        End Select
     End If
Next i
'------------Counting aging based on team, ticket type and priority-----------------

WS_CSS.Select
'values insertion to the respective cells in the dashboard
'----------Incident----------
Range("D14:G14").Value = Days0_1_INC
Range("D15:G15").Value = Days2_3_INC
Range("D16:G16").Value = Days4_5_INC
Range("D17:G17").Value = Days6_7_INC
Range("D18:G18").Value = Days8_14_INC
Range("D19:G19").Value = Days15_30_INC
Range("D20:G20").Value = Days31_60_INC
Range("D21:G21").Value = Days61_90_INC
Range("D22:G22").Value = Days_GT_90_INC
Range("D23:G23").Value = Days_Sum_INC

'--------Service Request--------
Range("I14:L14").Value = Days0_1_SRQ
Range("I15:L15").Value = Days2_3_SRQ
Range("I16:L16").Value = Days4_5_SRQ
Range("I17:L17").Value = Days6_7_SRQ
Range("I18:L18").Value = Days8_14_SRQ
Range("I19:L19").Value = Days15_30_SRQ
Range("I20:L20").Value = Days31_60_SRQ
Range("I21:L21").Value = Days61_90_SRQ
Range("I22:L22").Value = Days_GT_90_SRQ
Range("I23:L23").Value = Days_Sum_SRQ

'--------Problem Statement--------
Range("N14:Q14").Value = Days0_1_PRB
Range("N15:Q15").Value = Days2_3_PRB
Range("N16:Q16").Value = Days4_5_PRB
Range("N17:Q17").Value = Days6_7_PRB
Range("N18:Q18").Value = Days8_14_PRB
Range("N19:Q19").Value = Days15_30_PRB
Range("N20:Q20").Value = Days31_60_PRB
Range("N21:Q21").Value = Days61_90_PRB
Range("N22:Q22").Value = Days_GT_90_PRB
Range("N23:Q23").Value = Days_Sum_PRB

'----------Total Incident----------
Cells(14, 8).Formula = "=Sum(D14:G14)"
Cells(15, 8).Formula = "=Sum(D15:G15)"
Cells(16, 8).Formula = "=Sum(D16:G16)"
Cells(17, 8).Formula = "=Sum(D17:G17)"
Cells(18, 8).Formula = "=Sum(D18:G18)"
Cells(19, 8).Formula = "=Sum(D19:G19)"
Cells(20, 8).Formula = "=Sum(D20:G20)"
Cells(21, 8).Formula = "=Sum(D21:G21)"
Cells(22, 8).Formula = "=Sum(D22:G22)"

'--------Total Service Request--------
Cells(14, 13).Formula = "=Sum(I14:L14)"
Cells(15, 13).Formula = "=Sum(I15:L15)"
Cells(16, 13).Formula = "=Sum(I16:L16)"
Cells(17, 13).Formula = "=Sum(I17:L17)"
Cells(18, 13).Formula = "=Sum(I18:L18)"
Cells(19, 13).Formula = "=Sum(I19:L19)"
Cells(20, 13).Formula = "=Sum(I20:L20)"
Cells(21, 13).Formula = "=Sum(I21:L21)"
Cells(22, 13).Formula = "=Sum(I22:L22)"

'--------Total Problem Statement--------
Cells(14, 18).Formula = "=Sum(N14:Q14)"
Cells(15, 18).Formula = "=Sum(N15:Q15)"
Cells(16, 18).Formula = "=Sum(N16:Q16)"
Cells(17, 18).Formula = "=Sum(N17:Q17)"
Cells(18, 18).Formula = "=Sum(N18:Q18)"
Cells(19, 18).Formula = "=Sum(N19:Q19)"
Cells(20, 18).Formula = "=Sum(N20:Q20)"
Cells(21, 18).Formula = "=Sum(N21:Q21)"
Cells(22, 18).Formula = "=Sum(N22:Q22)"


End Sub
