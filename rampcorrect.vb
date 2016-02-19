Function GetDesBlockRow(ByVal Block As Integer) As Integer
    Dim DesRow As Integer
    If Block = 0 Then
        GetDesBlockRow = 110
    Else
        DesRow = Sheets("DESIRED").Cells(106, 2).Value
        GetDesBlockRow = DesRow + Block - 1
    End If
End Function

Function GetDCBlockRow(ByVal Block As Integer) As Integer
    Dim DCRow As Integer
    DCRow = Sheets("ONBARDC").Cells(104, 2).Value
    GetDCBlockRow = DCRow + Block - 1
End Function
Function AlphaColumn(ByVal Alphabet As String) As Integer
    AlphaColumn = Range(Alphabet & 1).Column
End Function
Sub dcFetch()
'
' dcFetch Macro
'
    Application.ScreenUpdating = False
    Dim i, j As Integer
    For i = 1 To 100
        For j = 1 To 32
            Sheets("ONBARDC").Cells(i, j).FormulaR1C1 = "=[DC_SHEET.xlsm]ONBARDC!R" & CStr(i) & "C" & CStr(j)
        Next j
    Next i
    Application.ScreenUpdating = True
End Sub
''''''''''''''''''''''''''''''''''''''''''''
'Main WorkFlow
''''''''''''''''''''''''''''''''''''''''''''
''Goto a row in feasible sheet
''Here everything is initially desired
''If the feasible sum of the row is violating the ramp
''''Ramp up/ramp down the prev row by distributing maximum ramp
Sub RampCorrect1()
''Get DC from other sheet
'First validate the headings of desired and entitlement and feasible
'Solve < 0 Constraint and > entitlement constraint
'Solve ramp till total available dissipation
'first make feasible = desiredfeasible
'In Desired Sheet
    Application.ScreenUpdating = False
    Dim i, j, k As Integer
    Dim DesStartCol, DesEndCol  As Integer
    DesStartCol = AlphaColumn(Sheets("DESIRED").Cells(103, 2).Value) '5
    DesEndCol = AlphaColumn(Sheets("DESIRED").Cells(104, 2).Value) '17
    
'In Entitlement Sheet
    Dim EntRow, EntStartCol, EntEndCol, DCCol  As Integer
    DCCol = AlphaColumn(Sheets("ONBARDC").Cells(103, 2).Value) '3
    EntRow = Sheets("ENTS").Cells(105, 2).Value ''2
    EntStartCol = AlphaColumn(Sheets("ENTS").Cells(103, 2).Value) '3
    EntEndCol = AlphaColumn(Sheets("ENTS").Cells(104, 2).Value) '15
    
    Dim DesSums(1 To 96) As Double
    Dim DesDiffs(1 To 96) As Double
    Dim Ent As Double
    
'Make feasible, desiredlegal = desired
    For i = 1 To 96 'here i is block
        For j = DesStartCol To DesEndCol
            Ent = Sheets("ENTS").Cells(GetDesBlockRow(i), j - DesStartCol + EntStartCol).Value * 0.01 * Sheets("ONBARDC").Cells(GetDCBlockRow(i), DCCol).Value
            If (Sheets("DESIRED").Cells(GetDesBlockRow(i), j).Value = "FULL" Or Sheets("DESIRED").Cells(GetDesBlockRow(i), j).Value = "full" Or Sheets("DESIRED").Cells(GetDesBlockRow(i), j).Value = "Full") Then
                Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Value = Ent
            End If
            Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Value = Sheets("DESIRED").Cells(GetDesBlockRow(i), j).Value
            
            If (Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j) < 0) Then
                Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j) = 0
            End If
            If Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Value > Ent Then
                Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Value = Ent
            End If
            Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value = Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Value
        Next j
    Next i
            
            
'Calculate DesiredSum for 1st row
    For i = 1 To 1
        DesSums(i) = 0
        For j = DesStartCol To DesEndCol
            DesSums(i) = DesSums(i) + Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value
        Next j
    Next i
    Dim MaxRampCol As Integer
    
    MaxRampCol = AlphaColumn(Sheets("DESIRED").Cells(105, 2).Value) '2
'MsgBox (Sheets("DESIRED").Cells(MaxRampRow, MaxRampCol))
    Dim MaxRamp As Double
    'Now solve ramps row by row from rows 2 to 96 rows

    For i = 2 To 96
    '''''''''''''We are in a bock Now
    
    'Calculate DesiredSums and DesiredDifferences
        DesSums(i) = 0
        For j = DesStartCol To DesEndCol
            DesSums(i) = DesSums(i) + Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value
        Next j
        DesDiffs(i) = DesSums(i) - DesSums(i - 1)
        'Color Row Total Sum Figure as White Before checking for Ramp Violation of the Row
        Sheets("DESIRED").Cells(GetDesBlockRow(i), DesEndCol + 4).Interior.Color = xlNone
        Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), DesEndCol + 4).Interior.Color = xlNone
        Sheets("FEASIBLE").Cells(GetDesBlockRow(i), DesEndCol + 4).Interior.Color = xlNone
        For j = DesStartCol To DesEndCol
            Sheets("DESIRED").Cells(GetDesBlockRow(i), j).Interior.Color = xlNone
            Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = xlNone
            Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = xlNone
        Next j
        If Abs(DesDiffs(i)) - 1 > Sheets("DESIRED").Cells(GetDesBlockRow(i), MaxRampCol).Value Then
        '''''''''''''We are in a Ramp Violating row Now
                Sheets("DESIRED").Cells(GetDesBlockRow(i), DesEndCol + 4).Interior.Color = RGB(255, 255, 0)
                Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), DesEndCol + 4).Interior.Color = RGB(255, 255, 0)
                Sheets("FEASIBLE").Cells(GetDesBlockRow(i), DesEndCol + 4).Interior.Color = RGB(255, 255, 0)
                
                ''Available rampup = something
                MaxRamp = Sheets("DESIRED").Cells(GetDesBlockRow(i), MaxRampCol).Value
                ''Desired Ramps of recipients  = array
                ''Entitlements of recipients  = array
                ''Given Ramps to recipients  = array initially zero
                Dim desiredRamps(1 To 50) As Double
                Dim cateredRamps(1 To 50) As Double
                Dim Entitlements(1 To 50) As Double
                If (DesDiffs(i) > 0) Then
                '''''''''''''We are in +ve Ramp Violating Row
                    
                    For j = DesStartCol To DesEndCol
                    '''''''''''''We are in +ve Row Column
                        desiredRamps(j) = Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value - Sheets("FEASIBLE").Cells(GetDesBlockRow(i - 1), j).Value
                        cateredRamps(j) = 0
                        Entitlements(j) = Sheets("ENTS").Cells(GetDesBlockRow(i), j - DesStartCol + EntStartCol).Value * 0.01 * Sheets("ONBARDC").Cells(GetDCBlockRow(i), DCCol).Value
                        ''Highlight the cell if desired Ramp is possitive
                        If desiredRamps(j) > 0 Then
                            Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = RGB(255, 255, 0)
                        Else
                            Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = xlNone
                        End If
                        
                    '''''''''''''We are out of +ve Row Column
                    Next j
                    
                    ''Calculate the ramps to be catered
                    DistributeRamp MaxRamp, DesStartCol, DesEndCol, desiredRamps:=desiredRamps, cateredRamps:=cateredRamps, Entitlements:=Entitlements
                    
                    ''Give the ramps that are calculated from the function to the cells
                    For j = DesStartCol To DesEndCol
                    '''''''''''''We are in +ve Row Column
                        Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value = Sheets("FEASIBLE").Cells(GetDesBlockRow(i - 1), j).Value + cateredRamps(j)
                    '''''''''''''We are out of +ve Row Column
                    Next j
                '''''''''''''We out of +ve Ramp Violating Row
                ElseIf (DesDiffs(i) < 0) Then
                '''''''''''''We are in -ve Ramp Violating Row
                    For j = DesStartCol To DesEndCol
                    '''''''''''''We are in -ve Row Column
                        desiredRamps(j) = -Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value + Sheets("FEASIBLE").Cells(GetDesBlockRow(i - 1), j).Value
                        cateredRamps(j) = 0
                        Entitlements(j) = Sheets("ENTS").Cells(GetDesBlockRow(i), j - DesStartCol + EntStartCol).Value * 0.01 * Sheets("ONBARDC").Cells(GetDCBlockRow(i), DCCol).Value
                    '''''''''''''We are out of -ve Row Column
                    Next j
                    
                    ''Calculate the ramps to be catered
                    DistributeRamp MaxRamp, DesStartCol, DesEndCol, desiredRamps:=desiredRamps, cateredRamps:=cateredRamps, Entitlements:=Entitlements
                    
                    ''Give the ramps that are calculated from the function to the cells
                    For j = DesStartCol To DesEndCol
                    '''''''''''''We are in +ve Row Column
                        Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value = Sheets("FEASIBLE").Cells(GetDesBlockRow(i - 1), j).Value - cateredRamps(j)
                    '''''''''''''We are out of -ve Row Column
                    Next j
                    
                
                '''''''''''''We out of -ve Ramp Violating Row
                End If
        '''''''''''''We out of a Ramp Violating row Now
        End If
        'Recalculating the sum of row
        DesSums(i) = 0
        For j = DesStartCol To DesEndCol
            DesSums(i) = DesSums(i) + Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value
        Next j
    '''''''''''''We got out of a block Now
    Next i
    Application.ScreenUpdating = True
End Sub

''''''''''''''''''''''''''''''''''''''''''''
'Function to distribute ramp
''''''''''''''''''''''''''''''''''''''''''''
''Available rampup = something
''Desired Ramps of recipients  = array
''Entitlements of recipients  = array
''Given Ramps to recipients  = array initially zero
''While Availablerampup > 1
''''''''Create a group of recipients whose given ramp < desired ramp
''''''''Distribute Available Rampup among the group recipients according to entitlment ratios and update Given Ramps
''''''''Update Available rampup
Function DistributeRamp(ByVal availableRamp As Double, ByVal DesStartCol As Integer, ByVal DesEndCol As Integer, ByRef desiredRamps() As Double, ByRef cateredRamps() As Double, ByRef Entitlements() As Double)
    Dim Iter As Integer
    Iter = 0
    Do While availableRamp > 1 And Iter < 100
        Iter = Iter + 1
        'Create a group of recipients whose given ramp < desired ramp
        Dim ViloateList(1 To 50) As Integer
        Dim NumberOfViolated As Integer
        Dim EntSum As Double
        NumberOfViolated = 0
        EntSum = 0
        Dim i As Integer
        For i = DesStartCol To DesEndCol
            If (cateredRamps(i) < desiredRamps(i) And desiredRamps(i) > 0) Then
                NumberOfViolated = NumberOfViolated + 1
                ViloateList(NumberOfViolated) = i
                'Update the Entitlement Sum
                EntSum = EntSum + Entitlements(i)
            End If
        Next i
        Dim Ramped, rampUp As Double
        Ramped = 0
        'Distribute Available Rampup among the group recipients according to entitlment ratios and update Given Ramps
        If EntSum > 0 Then
            For i = 1 To NumberOfViolated
                'Distribute Available Rampup according to entitlment ratios
                rampUp = availableRamp * Entitlements(ViloateList(i)) / EntSum
                If cateredRamps(ViloateList(i)) + rampUp > desiredRamps(ViloateList(i)) Then
                    rampUp = desiredRamps(ViloateList(i)) - cateredRamps(ViloateList(i))
                End If
                Ramped = Ramped + rampUp
                'Update Given Ramps
                cateredRamps(ViloateList(i)) = cateredRamps(ViloateList(i)) + rampUp
            Next i
        End If
        'Update Available rampup
        availableRamp = availableRamp - Ramped
    Loop
End Function
Sub RampCorrect2()
''Get DC from other sheet
'First validate the headings of desired and entitlement and feasible
'Solve < 0 Constraint and > entitlement constraint
'Solve ramp till total available dissipation
'first make feasible = desiredfeasible
'In Desired Sheet
    Application.ScreenUpdating = False
    Dim i, j, k As Integer
    Dim DesStartCol, DesEndCol  As Integer
    DesStartCol = AlphaColumn(Sheets("DESIRED").Cells(103, 2).Value) '5
    DesEndCol = AlphaColumn(Sheets("DESIRED").Cells(104, 2).Value) '17
    
'In Entitlement Sheet
    Dim EntRow, EntStartCol, EntEndCol, DCCol  As Integer
    DCCol = AlphaColumn(Sheets("ONBARDC").Cells(103, 2).Value) '3
    EntRow = Sheets("ENTS").Cells(105, 2).Value ''2
    EntStartCol = AlphaColumn(Sheets("ENTS").Cells(103, 2).Value) '3
    EntEndCol = AlphaColumn(Sheets("ENTS").Cells(104, 2).Value) '15
    
    Dim DesSums(1 To 96) As Double
    Dim DesDiffs(1 To 96) As Double
    Dim Ent As Double
    
    Dim MaxRampCol As Integer
    MaxRampCol = AlphaColumn(Sheets("DESIRED").Cells(105, 2).Value) '2
'MsgBox (Sheets("DESIRED").Cells(MaxRampRow, MaxRampCol))

    Dim MaxRamp As Double
    Dim desiredRamps(1 To 50) As Double
    Dim cateredRamps(1 To 50) As Double
    Dim Entitlements(1 To 50) As Double
    
'First check 110 row for valid numbers
    Dim ZeroTrue As Boolean
    ZeroTrue = True
    i = 0 'here i is block
    For j = DesStartCol To DesEndCol
        If IsEmpty(Sheets("DESIRED").Cells(GetDesBlockRow(i), j).Value) Then
            ZeroTrue = False
        End If
    Next j
    If ZeroTrue Then
    'Calculate all the shit
    'Make feasible, desiredlegal = desired for zero block
        i = 0
        For j = DesStartCol To DesEndCol
        'Assumption Yesterday 96 entitlement = today 1st entitlement and same for onbar dc
            Ent = Sheets("ENTS").Cells(GetDesBlockRow(1), j - DesStartCol + EntStartCol).Value * 0.01 * Sheets("ONBARDC").Cells(GetDCBlockRow(1), DCCol).Value
            If (Sheets("DESIRED").Cells(GetDesBlockRow(i), j).Value = "FULL" Or Sheets("DESIRED").Cells(GetDesBlockRow(i), j).Value = "full" Or Sheets("DESIRED").Cells(GetDesBlockRow(i), j).Value = "Full") Then
                Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Value = Ent
            End If
            Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Value = Sheets("DESIRED").Cells(GetDesBlockRow(i), j).Value
            
            If (Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j) < 0) Then
                Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j) = 0
            End If
            If Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Value > Ent Then
                Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Value = Ent
            End If
            Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value = Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Value
        Next j
        For j = DesStartCol To DesEndCol
            DesSumsZero = DesSumsZero + Sheets("FEASIBLE").Cells(GetDesBlockRow(0), j).Value
        Next j
        
    i = 1
    '''''''''''''We are in a bock Now
    
    'Calculate DesiredSums and DesiredDifferences
        DesSums(i) = 0
        For j = DesStartCol To DesEndCol
            DesSums(i) = DesSums(i) + Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value
        Next j
        DesDiffs(i) = DesSums(i) - DesSumsZero
        'Color Row Total Sum Figure as White Before checking for Ramp Violation of the Row
        Sheets("DESIRED").Cells(GetDesBlockRow(i), DesEndCol + 4).Interior.Color = xlNone
        Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), DesEndCol + 4).Interior.Color = xlNone
        Sheets("FEASIBLE").Cells(GetDesBlockRow(i), DesEndCol + 4).Interior.Color = xlNone
        For j = DesStartCol To DesEndCol
            Sheets("DESIRED").Cells(GetDesBlockRow(i), j).Interior.Color = xlNone
            Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = xlNone
            Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = xlNone
        Next j
        If Abs(DesDiffs(i)) - 1 > Sheets("DESIRED").Cells(GetDesBlockRow(i), MaxRampCol).Value Then
        '''''''''''''We are in a Ramp Violating row Now
                Sheets("DESIRED").Cells(GetDesBlockRow(i), DesEndCol + 4).Interior.Color = RGB(255, 255, 0)
                Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), DesEndCol + 4).Interior.Color = RGB(255, 255, 0)
                Sheets("FEASIBLE").Cells(GetDesBlockRow(i), DesEndCol + 4).Interior.Color = RGB(255, 255, 0)
                ''Available rampup = something
                MaxRamp = Sheets("DESIRED").Cells(GetDesBlockRow(i), MaxRampCol).Value
                ''Desired Ramps of recipients  = array
                ''Entitlements of recipients  = array
                ''Given Ramps to recipients  = array initially zero
                If (DesDiffs(i) > 0) Then
                '''''''''''''We are in +ve Ramp Violating Row
                    
                    For j = DesStartCol To DesEndCol
                    '''''''''''''We are in +ve Row Column
                        desiredRamps(j) = Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value - Sheets("FEASIBLE").Cells(GetDesBlockRow(i - 1), j).Value
                        cateredRamps(j) = 0
                        Entitlements(j) = Sheets("ENTS").Cells(GetDesBlockRow(1), j - DesStartCol + EntStartCol).Value * 0.01 * Sheets("ONBARDC").Cells(GetDCBlockRow(1), DCCol).Value
                        ''Highlight the cell if desired Ramp is possitive
                        If desiredRamps(j) > 0 Then
                            Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = RGB(255, 255, 0)
                            Sheets("DESIRED").Cells(GetDesBlockRow(i), j).Interior.Color = RGB(255, 255, 0)
                            Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = RGB(255, 255, 0)
                        Else
                            Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = xlNone
                            Sheets("DESIRED").Cells(GetDesBlockRow(i), j).Interior.Color = xlNone
                            Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = xlNone
                        End If
                        
                    '''''''''''''We are out of +ve Row Column
                    Next j
                    
                    ''Calculate the ramps to be catered
                    DistributeRamp MaxRamp, DesStartCol, DesEndCol, desiredRamps:=desiredRamps, cateredRamps:=cateredRamps, Entitlements:=Entitlements
                    
                    ''Give the ramps that are calculated from the function to the cells
                    For j = DesStartCol To DesEndCol
                    '''''''''''''We are in +ve Row Column
                        Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value = Sheets("FEASIBLE").Cells(GetDesBlockRow(i - 1), j).Value + cateredRamps(j)
                    '''''''''''''We are out of +ve Row Column
                    Next j
                '''''''''''''We out of +ve Ramp Violating Row
                ElseIf (DesDiffs(i) < 0) Then
                '''''''''''''We are in -ve Ramp Violating Row
                    For j = DesStartCol To DesEndCol
                    '''''''''''''We are in -ve Row Column
                        desiredRamps(j) = -Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value + Sheets("FEASIBLE").Cells(GetDesBlockRow(i - 1), j).Value
                        cateredRamps(j) = 0
                        Entitlements(j) = Sheets("ENTS").Cells(GetDesBlockRow(1), j - DesStartCol + EntStartCol).Value * 0.01 * Sheets("ONBARDC").Cells(GetDCBlockRow(1), DCCol).Value
                        ''Highlight the cell if desired Ramp is possitive
                        If desiredRamps(j) > 0 Then
                            Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = RGB(0, 255, 255)
                            Sheets("DESIRED").Cells(GetDesBlockRow(i), j).Interior.Color = RGB(0, 255, 255)
                            Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = RGB(0, 255, 255)
                        Else
                            Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = xlNone
                            Sheets("DESIRED").Cells(GetDesBlockRow(i), j).Interior.Color = xlNone
                            Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = xlNone
                        End If
                    '''''''''''''We are out of -ve Row Column
                    Next j
                    
                    ''Calculate the ramps to be catered
                    DistributeRamp MaxRamp, DesStartCol, DesEndCol, desiredRamps:=desiredRamps, cateredRamps:=cateredRamps, Entitlements:=Entitlements
                    
                    ''Give the ramps that are calculated from the function to the cells
                    For j = DesStartCol To DesEndCol
                    '''''''''''''We are in +ve Row Column
                        Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value = Sheets("FEASIBLE").Cells(GetDesBlockRow(i - 1), j).Value - cateredRamps(j)
                    '''''''''''''We are out of -ve Row Column
                    Next j
                    
                
                '''''''''''''We out of -ve Ramp Violating Row
                End If
        '''''''''''''We out of a Ramp Violating row Now
        End If
        'Recalculating the sum of row
        DesSums(i) = 0
        For j = DesStartCol To DesEndCol
            DesSums(i) = DesSums(i) + Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value
        Next j
        
    Else 'If ZeroTrue Then
        For j = DesStartCol To DesEndCol
            Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value = "NA"
            Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value = "NA"
            Sheets("FEASIBLE").Cells(GetDesBlockRow(1), j).Interior.ColorIndex = xlNone
            Sheets("FEASIBLE").Cells(GetDesBlockRow(1), j).Interior.ColorIndex = xlNone
        Next j
    End If 'If ZeroTrue Then
    
'Make feasible, desiredlegal = desired
    Dim Start As Integer
    Start = 1
    If ZeroTrue Then
        Start = 2
    End If
    For i = Start To 96 'here i is block
        For j = DesStartCol To DesEndCol
            Ent = Sheets("ENTS").Cells(GetDesBlockRow(i), j - DesStartCol + EntStartCol).Value * 0.01 * Sheets("ONBARDC").Cells(GetDCBlockRow(i), DCCol).Value
            If (Sheets("DESIRED").Cells(GetDesBlockRow(i), j).Value = "FULL" Or Sheets("DESIRED").Cells(GetDesBlockRow(i), j).Value = "full" Or Sheets("DESIRED").Cells(GetDesBlockRow(i), j).Value = "Full") Then
                Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Value = Ent
            End If
            Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Value = Sheets("DESIRED").Cells(GetDesBlockRow(i), j).Value
            
            If (Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j) < 0) Then
                Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j) = 0
            End If
            If Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Value > Ent Then
                Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Value = Ent
            End If
            Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value = Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Value
        Next j
    Next i
            
            
'Calculate DesiredSum for 1st row
    For i = 1 To 1
        DesSums(i) = 0
        For j = DesStartCol To DesEndCol
            DesSums(i) = DesSums(i) + Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value
        Next j
    Next i
    
    'Now solve ramps row by row from rows 2 to 96 rows
    For i = 2 To 96
    '''''''''''''We are in a bock Now
    
    'Calculate DesiredSums and DesiredDifferences
        DesSums(i) = 0
        For j = DesStartCol To DesEndCol
            DesSums(i) = DesSums(i) + Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value
        Next j
        DesDiffs(i) = DesSums(i) - DesSums(i - 1)
        'Color Row Total Sum Figure as White Before checking for Ramp Violation of the Row
        Sheets("DESIRED").Cells(GetDesBlockRow(i), DesEndCol + 4).Interior.Color = xlNone
        Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), DesEndCol + 4).Interior.Color = xlNone
        Sheets("FEASIBLE").Cells(GetDesBlockRow(i), DesEndCol + 4).Interior.Color = xlNone
        For j = DesStartCol To DesEndCol
            Sheets("DESIRED").Cells(GetDesBlockRow(i), j).Interior.Color = xlNone
            Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = xlNone
            Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = xlNone
        Next j
        If Abs(DesDiffs(i)) - 1 > Sheets("DESIRED").Cells(GetDesBlockRow(i), MaxRampCol).Value Then
        '''''''''''''We are in a Ramp Violating row Now
                Sheets("DESIRED").Cells(GetDesBlockRow(i), DesEndCol + 4).Interior.Color = RGB(255, 255, 0)
                Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), DesEndCol + 4).Interior.Color = RGB(255, 255, 0)
                Sheets("FEASIBLE").Cells(GetDesBlockRow(i), DesEndCol + 4).Interior.Color = RGB(255, 255, 0)
                
                ''Available rampup = something
                MaxRamp = Sheets("DESIRED").Cells(GetDesBlockRow(i), MaxRampCol).Value
                ''Desired Ramps of recipients  = array
                ''Entitlements of recipients  = array
                ''Given Ramps to recipients  = array initially zero
                If (DesDiffs(i) > 0) Then
                '''''''''''''We are in +ve Ramp Violating Row
                    
                    For j = DesStartCol To DesEndCol
                    '''''''''''''We are in +ve Row Column
                        desiredRamps(j) = Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value - Sheets("FEASIBLE").Cells(GetDesBlockRow(i - 1), j).Value
                        cateredRamps(j) = 0
                        Entitlements(j) = Sheets("ENTS").Cells(GetDesBlockRow(i), j - DesStartCol + EntStartCol).Value * 0.01 * Sheets("ONBARDC").Cells(GetDCBlockRow(i), DCCol).Value
                        ''Highlight the cell if desired Ramp is possitive
                        If desiredRamps(j) > 0 Then
                            Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = RGB(255, 255, 0)
                            Sheets("DESIRED").Cells(GetDesBlockRow(i), j).Interior.Color = RGB(255, 255, 0)
                            Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = RGB(255, 255, 0)
                        Else
                            Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = xlNone
                            Sheets("DESIRED").Cells(GetDesBlockRow(i), j).Interior.Color = xlNone
                            Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = xlNone
                        End If
                        
                    '''''''''''''We are out of +ve Row Column
                    Next j
                    
                    ''Calculate the ramps to be catered
                    DistributeRamp MaxRamp, DesStartCol, DesEndCol, desiredRamps:=desiredRamps, cateredRamps:=cateredRamps, Entitlements:=Entitlements
                    
                    ''Give the ramps that are calculated from the function to the cells
                    For j = DesStartCol To DesEndCol
                    '''''''''''''We are in +ve Row Column
                        Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value = Sheets("FEASIBLE").Cells(GetDesBlockRow(i - 1), j).Value + cateredRamps(j)
                    '''''''''''''We are out of +ve Row Column
                    Next j
                '''''''''''''We out of +ve Ramp Violating Row
                ElseIf (DesDiffs(i) < 0) Then
                '''''''''''''We are in -ve Ramp Violating Row
                    For j = DesStartCol To DesEndCol
                    '''''''''''''We are in -ve Row Column
                        desiredRamps(j) = -Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value + Sheets("FEASIBLE").Cells(GetDesBlockRow(i - 1), j).Value
                        cateredRamps(j) = 0
                        Entitlements(j) = Sheets("ENTS").Cells(GetDesBlockRow(i), j - DesStartCol + EntStartCol).Value * 0.01 * Sheets("ONBARDC").Cells(GetDCBlockRow(i), DCCol).Value
                        ''Highlight the cell if desired Ramp is possitive
                        If desiredRamps(j) > 0 Then
                            Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = RGB(0, 255, 255)
                            Sheets("DESIRED").Cells(GetDesBlockRow(i), j).Interior.Color = RGB(0, 255, 255)
                            Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = RGB(0, 255, 255)
                        Else
                            Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = xlNone
                            Sheets("DESIRED").Cells(GetDesBlockRow(i), j).Interior.Color = xlNone
                            Sheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = xlNone
                        End If
                    '''''''''''''We are out of -ve Row Column
                    Next j
                    
                    ''Calculate the ramps to be catered
                    DistributeRamp MaxRamp, DesStartCol, DesEndCol, desiredRamps:=desiredRamps, cateredRamps:=cateredRamps, Entitlements:=Entitlements
                    
                    ''Give the ramps that are calculated from the function to the cells
                    For j = DesStartCol To DesEndCol
                    '''''''''''''We are in +ve Row Column
                        Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value = Sheets("FEASIBLE").Cells(GetDesBlockRow(i - 1), j).Value - cateredRamps(j)
                    '''''''''''''We are out of -ve Row Column
                    Next j
                    
                
                '''''''''''''We out of -ve Ramp Violating Row
                End If
        '''''''''''''We out of a Ramp Violating row Now
        End If
        'Recalculating the sum of row
        DesSums(i) = 0
        For j = DesStartCol To DesEndCol
            DesSums(i) = DesSums(i) + Sheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value
        Next j
    '''''''''''''We got out of a block Now
    Next i
    
    ''Select the range of cells to be copied
    Sheets("FEASIBLE").Range(Cells(GetDesBlockRow(1), DesStartCol - 1), Cells(GetDesBlockRow(96), DesEndCol)).Select
    Application.ScreenUpdating = True
End Sub
