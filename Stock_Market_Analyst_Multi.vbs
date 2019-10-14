Sub Stock_Market_Analyst_Multi()

For Each ws In Worksheets
    'Find last row in a worksheet
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Keep tack of location for a summary table
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
    
    'Print column names for header in the summary table
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    
    'Define and initialize row count
    Dim Row_Count As Long
    Row_Count = 0
    
    
    For i = 2 To lastrow
    
        'Check if Ticker is still the same
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Write Ticker to summary table
            ws.Range("I" & Summary_Table_Row) = ws.Cells(i, 1).Value
            
            'Write Yearly change to Summary table and format color
            Change = ws.Cells(i, 6) - ws.Cells(i - Row_Count, 3)
            ws.Range("J" & Summary_Table_Row) = Change
            
            If Change < 0 Then
                'Negative is red
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            Else
                'Positive is green
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            End If
            
            
            'Write percent change to summary table
            If ws.Cells(i - Row_Count, 3) <> 0 Then
                ws.Range("K" & Summary_Table_Row) = Change / ws.Cells(i - Row_Count, 3)
            Else
				'If beginning value is zero display n/a
                ws.Range("K" & Summary_Table_Row) = "N/A"
            End If
            
            'Write total volume to summary table
            ws.Range("L" & Summary_Table_Row) = WorksheetFunction.Sum(ws.Range(ws.Cells(i - Row_Count, 7), ws.Cells(i, 7)))
            'Range("L" & Summary_Table_Row).Formula = "=SUM(" & Range(Cells(i - Row_Count, 7), Cells(i, 7)).Address(False, False) & ")"
            
            'Increment summary table row and reset row count
            Summary_Table_Row = Summary_Table_Row + 1
            Row_Count = 0
            
        Else
        
            'Increment row count
            Row_Count = Row_Count + 1
            
        End If
    Next i
    
    'Format the Percent Change column in the summary table
    '(Summary_Table_Row is 1 too large after exiting the for loop)
    ws.Range("K2:K" & Summary_Table_Row - 1).NumberFormat = "0.00%"
    
    'Make Second Summary Table
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    
    'Range("Q3").Formula = "=MIN(" & Range("K2:K" & Summary_Table_Row - 1).Address(False, False) & ")"
    
    'Greatest Percent Increase
    ws.Range("O2") = "Greatest % Increase"
    max_p = WorksheetFunction.Max(ws.Range("K2:K" & Summary_Table_Row - 1))
    ws.Range("Q2") = max_p
    loc_max_p = WorksheetFunction.Match(max_p, ws.Range("K2:K" & Summary_Table_Row - 1), 0)
    ws.Range("P2") = WorksheetFunction.Index(ws.Range("I2:I" & Summary_Table_Row - 1), loc_max_p, 0)
    ws.Range("Q2").NumberFormat = "0.00%"
    
    'Greatest Percent Decrease
    ws.Range("O3") = "Greatest % Decrease"
    min_p = WorksheetFunction.Min(ws.Range("K2:K" & Summary_Table_Row - 1))
    ws.Range("Q3") = min_p
    loc_min_p = WorksheetFunction.Match(min_p, ws.Range("K2:K" & Summary_Table_Row - 1), 0)
    ws.Range("P3") = WorksheetFunction.Index(ws.Range("I2:I" & Summary_Table_Row - 1), loc_min_p, 0)
    ws.Range("Q3").NumberFormat = "0.00%"
    
    'Greatest Total Volume
    ws.Range("O4") = "Greatest Total Volume"
    max_v = WorksheetFunction.Max(ws.Range("L2:L" & Summary_Table_Row - 1))
    ws.Range("Q4") = max_v
    loc_max_v = WorksheetFunction.Match(max_v, ws.Range("L2:L" & Summary_Table_Row - 1), 0)
    ws.Range("P4") = WorksheetFunction.Index(ws.Range("I2:I" & Summary_Table_Row - 1), loc_max_v, 0)
    
    'Increase Column width to see entire headers
    ws.Columns("J:J").EntireColumn.AutoFit
    ws.Columns("K:K").EntireColumn.AutoFit
    ws.Columns("L:L").EntireColumn.AutoFit
    ws.Columns("O:O").EntireColumn.AutoFit
    
Next ws

End Sub
