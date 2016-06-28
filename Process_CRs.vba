Private Sub Process_CRs_Click()

    ' Optimization Attempt
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False

        Dim bp_cr As Worksheet
        Dim i As Integer, j As Integer, h_i As Integer 

        ' Process selected sheets
        For Each bp_cr In Worksheets
            Select Case bp_cr.CodeName
                Case "Sheet5", "Sheet8", "Sheet111", "Sheet14"
                With bp_cr
                
                    For h_i = 1 To 2 
                        i = 2 ' Skip Headers
                
                        ' Iterate through the records capturing ID, and WR info
                        Do Until IsEmpty(.Cells(i, 1))
                            cur_id = .Cells(i, 1)    ' 1 = CR_ID
                            cur_type = .Cells(i, 12) ' 12 = WR_type
                
                            ' If key found, backtrack to acertain removal criteria satisfied
                            If cur_type = "Solution" And i > 2 Then
                                 j = i - 1
                                 prev_id = .Cells(j, 1)
                                 prev_type = .Cells(j, 12)
                                    Do While (prev_id = cur_id And prev_type <> "Solution")
                                        .Rows(j).EntireRow.Delete
                                        j = j - 1
                                        prev_id = .Cells(j, 1)
                                        prev_type = .Cells(j, 12)
                                    Loop
                            End If
                            i = i + 1
                        Loop
                    Next h_i 
                    
                    ' Remove the dups on CR_ID criteria
                    .Cells.RemoveDuplicates Columns:=Array(1) 
                    
                End With
            End Select
        Next bp_cr
        
        ' Return Application To Initial State
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        
    End With
End Sub
