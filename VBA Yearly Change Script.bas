Attribute VB_Name = "Module2"
Sub Yearly_Change():

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

First_Open_Price = Cells(2, 3)
Dim Yearly_Change As Double
Dim Open_Price As Double
Dim Close_Price As Double
Dim i As Integer


'For i = 2 To 797711
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Close_Price = Cells(i, 6).Value
        MsgBox (Close_Price)
        'Open_Price = Cells(i + 1, 3).Value
        'MsgBox (Open_Price)
        'Yearly_Change = Open_Price - Close_Price
        'Range("J" & Summary_Table_Row).Value = Yearly_Change
    End If
    
'Next i

End Sub

