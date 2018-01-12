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


Sheets("MainData").Range("A1:AA1").Select
Selection.AutoFilter

With Selection
    
     '.AutoFilter Field:=25, Criteria1:="", Operator:=xlOr, Criteria2:=">=" & CLng(DateOfreport)
     '------------------ Filtering Data for TEAM ----------------------
     .AutoFilter Field:=8, Criteria1:=team
     
     For i = 2 To DAlro
     
        If Cells(i, 25).Value = "" Or Cells(i, 25).Value >= CLng(DateOfreport) Then
            If Cells(i, 24).Value <> "" And Cells(i, 24).Value <= CLng(DateOfreport) Then
                Cells(i, 19).Value = CLng(DateOfreport) - Cells(i, 24).Value
            Else
                If Cells(i, 24).Value = "" And Cells(i, 23).Value <= CLng(DateOfreport) Then
                    Cells(i, 19).Value = CLng(DateOfreport) - Cells(i, 23).Value
                End If
            End If
        End If
        
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
                                    Days_Sum_INC(0) = Days_Sum_INC(0) + 1
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
                                    Days_Sum_INC(1) = Days_Sum_INC(1) + 1
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
                                    Days_Sum_INC(2) = Days_Sum_INC(2) + 1
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
                                    Days_Sum_INC(3) = Days_Sum_INC(3) + 1
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
                                    Days_Sum_SRQ(0) = Days_Sum_SRQ(0) + 1
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
                                    Days_Sum_SRQ(1) = Days_Sum_SRQ(1) + 1
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
                                    Days_Sum_SRQ(2) = Days_Sum_SRQ(2) + 1
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
                                    Days_Sum_SRQ(3) = Days_Sum_SRQ(3) + 1
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
                                    Days_Sum_PRB(0) = Days_Sum_PRB(0) + 1
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
                                    Days_Sum_PRB(1) = Days_Sum_PRB(1) + 1
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
                                    Days_Sum_PRB(2) = Days_Sum_PRB(2) + 1
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
                                    Days_Sum_PRB(3) = Days_Sum_PRB(3) + 1
                                End If
                            End If
                End Select
        End Select
     
     Next i
     
        .AutoFilter Field:=1, Criteria1:="INC"
        
            .AutoFilter Field:=12, Criteria1:="1"
            
            'Days0_1_INC (0)=

End With


'values insertion to the respective cells in the dashboard
'----------Incident----------
Range("D14:H14").Value = Days0_1_INC
Range("D15:H15").Value = Days2_3_INC
Range("D16:H16").Value = Days4_5_INC
Range("D17:H17").Value = Days6_7_INC
Range("D18:H18").Value = Days8_14_INC
Range("D19:H19").Value = Days15_30_INC
Range("D20:H20").Value = Days31_60_INC
Range("D21:H21").Value = Days61_90_INC
Range("D22:H22").Value = Days_GT_90_INC
Range("D23:H23").Value = Days_Sum_INC

'--------Service Request--------
Range("I14:M14").Value = Days0_1_SRQ
Range("I15:M15").Value = Days2_3_SRQ
Range("I16:M16").Value = Days4_5_SRQ
Range("I17:M17").Value = Days6_7_SRQ
Range("I18:M18").Value = Days8_14_SRQ
Range("I19:M19").Value = Days15_30_SRQ
Range("I20:M20").Value = Days31_60_SRQ
Range("I21:M21").Value = Days61_90_SRQ
Range("I22:M22").Value = Days_GT_90_SRQ
Range("I23:M23").Value = Days_Sum_SRQ

'--------Problem Statement--------
Range("N14:R14").Value = Days0_1_PRB
Range("N15:R15").Value = Days2_3_PRB
Range("N16:R16").Value = Days4_5_PRB
Range("N17:R17").Value = Days6_7_PRB
Range("N18:R18").Value = Days8_14_PRB
Range("N19:R19").Value = Days15_30_PRB
Range("N20:R20").Value = Days31_60_PRB
Range("N21:R21").Value = Days61_90_PRB
Range("N22:R22").Value = Days_GT_90_PRB
Range("N23:R23").Value = Days_Sum_PRB

End Sub
