Function GetDesBlockRow(ByVal Block As Integer) As Integer
    Dim DesRow As Integer
    DesRow = 2
    GetDesBlockRow = DesRow + Block - 1
End Function

Function GetDCBlockRow(ByVal Block As Integer) As Integer
    Dim DCRow As Integer
    DCRow = 3
    GetDCBlockRow = DCRow + Block - 1
End Function

''''''''''''''''''''''''''''''''''''''''''''
'Main WorkFlow
''''''''''''''''''''''''''''''''''''''''''''
''Goto a row in feasible sheet
''Here everything is initially desired
''If the feasible sum of the row is violating the ramp
''''Ramp up/ramp down the prev row by distributing ramp
Sub RampCorrect1()
'First validate the headings of desired and entitlement and feasible
'Solve < 0 Constraint and > entitlement constraint
'Solve ramp till total available dissipation
'first make feasible = desiredfeasible
    ''TODO Find the columns that can be ramped
'In Desired Sheet
    Dim DesStartCol, DesEndCol  As Integer
    DesStartCol = 5
    DesEndCol = 17
    
'In Entitlement Sheet
    Dim EntRow, EntStartCol, EntEndCol, DCCol  As Integer
    DCCol = 3
    EntRow = 2
    EntStartCol = 3
    EntEndCol = 15
    
    Dim DesSums(1 To 96) As Double
    Dim DesDiffs(1 To 96) As Double
    Dim i, j, k As Integer
    Dim Ent As Double
    
'Make feasible, desiredlegal = desired
    For i = 1 To 96
        For j = DesStartCol To DesEndCol
            Worksheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Value = Worksheets("DESIRED").Cells(GetDesBlockRow(i), j).Value
            Ent = Worksheets("ENTS").Cells(GetDesBlockRow(i), j - DesStartCol + EntStartCol).Value * 0.01 * Worksheets("DESIREDFEASIBLE").Cells(GetDCBlockRow(i), DCCol).Value
            If (Worksheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j) < 0) Then
                Worksheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j) = 0
            End If
            If Worksheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Value > Ent Then
                Worksheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Value = Ent
            End If
            Worksheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value = Worksheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Value
        Next j
    Next i
            
            
'Calculate DesiredSum for 1st row
    For i = 1 To 1
        DesSums(i) = 0
        For j = DesStartCol To DesEndCol
            DesSums(i) = DesSums(i) + Worksheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), j).Value
        Next j
    Next i
    Dim MaxRampRow, MaxRampCol As Integer
    MaxRampRow = 2
    MaxRampCol = 2
'MsgBox (Worksheets("DESIRED").Cells(MaxRampRow, MaxRampCol))
    Dim MaxRamp As Double
    'Now solve ramps row by row from rows 2 to 96 rows

    For i = 2 To 96
    '''''''''''''We are in a bock Now
    
    'Calculate DesiredSums and DesiredDifferences
        DesSums(i) = 0
        For j = DesStartCol To DesEndCol
            DesSums(i) = DesSums(i) + Worksheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value
        Next j
        DesDiffs(i) = DesSums(i) - DesSums(i - 1)
        'Color Row Total Sum Figure as White Before checking for Ramp Violation of the Row
        Worksheets("DESIRED").Cells(GetDesBlockRow(i), DesEndCol + 4).Interior.Color = RGB(255, 255, 255)
        Worksheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), DesEndCol + 4).Interior.Color = RGB(255, 255, 255)
        Worksheets("FEASIBLE").Cells(GetDesBlockRow(i), DesEndCol + 4).Interior.Color = RGB(255, 255, 255)
        
        If Abs(DesDiffs(i)) - 1 > Worksheets("FEASIBLE").Cells(MaxRampRow + i - 1, MaxRampCol).Value Then
        '''''''''''''We are in a Ramp Violating row Now
                Worksheets("DESIRED").Cells(GetDesBlockRow(i), DesEndCol + 4).Interior.Color = RGB(255, 255, 0)
                Worksheets("DESIREDFEASIBLE").Cells(GetDesBlockRow(i), DesEndCol + 4).Interior.Color = RGB(255, 255, 0)
                Worksheets("FEASIBLE").Cells(GetDesBlockRow(i), DesEndCol + 4).Interior.Color = RGB(255, 255, 0)
                
                ''Available rampup = something
                MaxRamp = Worksheets("DESIRED").Cells(MaxRampRow + i - 1, MaxRampCol).Value
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
                        desiredRamps(j) = Worksheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value - Worksheets("FEASIBLE").Cells(GetDesBlockRow(i - 1), j).Value
                        cateredRamps(j) = 0
                        Entitlements(j) = Worksheets("ENTS").Cells(GetDesBlockRow(i), j - DesStartCol + EntStartCol).Value * 0.01 * Worksheets("DESIREDFEASIBLE").Cells(GetDCBlockRow(i), DCCol).Value
                        ''Highlight the cell if desired Ramp is possitive
                        If desiredRamps(j) > 0 Then
                            Worksheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = RGB(255, 255, 0)
                        Else
                            Worksheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Interior.Color = RGB(255, 255, 255)
                        End If
                        
                    '''''''''''''We are out of +ve Row Column
                    Next j
                    
                    ''Calculate the ramps to be catered
                    DistributeRamp MaxRamp, DesStartCol, DesEndCol, desiredRamps:=desiredRamps, cateredRamps:=cateredRamps, Entitlements:=Entitlements
                    
                    ''Give the ramps that are calculated from the function to the cells
                    For j = DesStartCol To DesEndCol
                    '''''''''''''We are in +ve Row Column
                        Worksheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value = Worksheets("FEASIBLE").Cells(GetDesBlockRow(i - 1), j).Value + cateredRamps(j)
                    '''''''''''''We are out of +ve Row Column
                    Next j
                '''''''''''''We out of +ve Ramp Violating Row
                ElseIf (DesDiffs(i) < 0) Then
                '''''''''''''We are in -ve Ramp Violating Row
                    For j = DesStartCol To DesEndCol
                    '''''''''''''We are in -ve Row Column
                        desiredRamps(j) = -Worksheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value + Worksheets("FEASIBLE").Cells(GetDesBlockRow(i - 1), j).Value
                        cateredRamps(j) = 0
                        Entitlements(j) = Worksheets("ENTS").Cells(GetDesBlockRow(i), j - DesStartCol + EntStartCol).Value * 0.01 * Worksheets("DESIREDFEASIBLE").Cells(GetDCBlockRow(i), DCCol).Value
                    '''''''''''''We are out of -ve Row Column
                    Next j
                    
                    ''Calculate the ramps to be catered
                    DistributeRamp MaxRamp, DesStartCol, DesEndCol, desiredRamps:=desiredRamps, cateredRamps:=cateredRamps, Entitlements:=Entitlements
                    
                    ''Give the ramps that are calculated from the function to the cells
                    For j = DesStartCol To DesEndCol
                    '''''''''''''We are in +ve Row Column
                        Worksheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value = Worksheets("FEASIBLE").Cells(GetDesBlockRow(i - 1), j).Value - cateredRamps(j)
                    '''''''''''''We are out of -ve Row Column
                    Next j
                    
                
                '''''''''''''We out of -ve Ramp Violating Row
                End If
        '''''''''''''We out of a Ramp Violating row Now
        End If
        'Recalculating the sum of row
        DesSums(i) = 0
        For j = DesStartCol To DesEndCol
            DesSums(i) = DesSums(i) + Worksheets("FEASIBLE").Cells(GetDesBlockRow(i), j).Value
        Next j
    '''''''''''''We got out of a block Now
    Next i
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
