Attribute VB_Name = "Module1"
Sub VBA_Stock__Complete():

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'##Begin For Loop
For i = 2 To 797711
    '## If Statement for Total_Stock_Volume and Ticker + Adding to Summary_Table
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        Ticker_Symbol = Cells(i, 1).Value
        Range("I" & Summary_Table_Row).Value = Ticker_Symbol

       Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
       Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
    
         '##Add conditional If statement for coloring
         'If Yearly_Change >= 0 Then
         '    Yearly_Change.Interior.ColorIndex = 4
         'Else
         '    Yearly_Change.Interior.ColorIndex = 3
         'End If

      Summary_Table_Row = Summary_Table_Row + 1
        Total_Stock_Volume = 0
    Else
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7)
    End If
    '## End If Statement for Total_Stock_Volume and Ticker + Adding to Summary_Table

    
            
Next i
End Sub

