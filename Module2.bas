Attribute VB_Name = "Module2"
Sub Do_Something():
Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call Stock_Analysis
        Call Conditional
    Next
    Application.ScreenUpdating = True
End Sub
Sub Stock_Analysis():
Dim Ticker As String
Dim BegValue As Double
Dim EndValue As Double
Dim SumVolume As Double

Dim row, row_amount, newRow As Integer

LastRow = Range("A" & Rows.Count).End(xlUp).row
newRow = 2


    For row = 2 To LastRow
    
        If row = 2 Then
    
            Ticker = Cells(row, 1).Value
            BegValue = Cells(row, 3).Value
            EndValue = Cells(row, 6).Value
            SumVolume = Cells(row, 7).Value
    
        Else
        
            If Cells(row, 1) = Ticker Then
            
                SumVolume = SumVolume + Cells(row, 7).Value
                EndValue = Cells(row, 6).Value
                
            Else
            
                Cells(newRow, 9).Value = Ticker
                Cells(newRow, 10).Value = EndValue - BegValue
                Cells(newRow, 11).Value = 100 * Cells(newRow, 10).Value / BegValue
                Cells(newRow, 12).Value = SumVolume
                newRow = newRow + 1
                
                Ticker = Cells(row, 1).Value
                BegValue = Cells(row, 3).Value
                EndValue = Cells(row, 6).Value
                SumVolume = Cells(row, 7).Value
                
            End If
                
        End If
        
    Next row
    
End Sub

Sub Conditional():
Dim i As Integer
GreatInc = Cells(2, 11).Value
GreatDec = Cells(2, 11).Value
GreatVol = Cells(2, 12).Value
Length = Range("J" & Rows.Count).End(xlUp).row
For i = 2 To Length
    If Cells(i, 10).Value < 0 Then
        Cells(i, 10).Interior.ColorIndex = 3
    Else
        Cells(i, 10).Interior.ColorIndex = 4
    End If
    
    
    If Cells(i, 11).Value > GreatInc Then
        GreatInc = Cells(i, 11).Value
        Cells(2, 16).Value = Cells(i, 9).Value
    Else
        GreatInc = GreatInc
    End If
    
    If Cells(i, 11).Value < GreatDec Then
        GreatDec = Cells(i, 11).Value
        Cells(3, 16).Value = Cells(i, 9).Value
    Else
        GreatDec = GreatDec
    End If
    
    If Cells(i, 12).Value > GreatVol Then
        GreatVol = Cells(i, 12).Value
        Cells(4, 16).Value = Cells(i, 9).Value
    Else
        GreatVol = GreatVol
    End If
    
    Cells(2, 17).Value = GreatInc / 100
    Cells(3, 17).Value = GreatDec / 100
    Cells(4, 17).Value = Format(GreatVol, "Scientific")
Next i
End Sub
