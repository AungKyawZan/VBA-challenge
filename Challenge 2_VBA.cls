VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stockdata():
    
Dim ticker As String
Dim yearly_change As Double
    

Dim volume As Double
Dim result As Integer
    
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

volume = 0
result = 2
p = 2


For o = 2 To 753001

If Cells(o + 1, 1).Value <> Cells(o, 1).Value Then
'ticker
    ticker = Cells(o, 1).Value
    Range("I" & result).Value = ticker
'yearly change
    
    Cells(result, 10).Value = Cells(o, 6).Value - Cells(p, 3).Value
    If Range("J" & result).Value < 0 Then
        Range("J" & result).Interior.ColorIndex = 3
    Else
        Range("J" & result).Interior.ColorIndex = 4
    End If
    
'% change
    Cells(result, 11).Value = ((Cells(o, 6).Value - Cells(p, 3).Value) / Cells(p, 3).Value)
    p = o + 1
'volume
    volume = volume + Cells(o, 7).Value

    Range("L" & result).Value = volume
'+1 row
    result = result + 1
'reset
    volume = 0
    
    Else
      volume = volume + Cells(o, 7).Value

End If

Next o
    
'BONUS
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % increase"
Cells(3, 15).Value = "Greatest % decrease"
Cells(4, 15).Value = "Greatest total volume"
   

'value
 Cells(2, 17).Value = val_increase
 Cells(3, 17).Value = val_decrease
 Cells(4, 17).Value = val_vol

p_change = Cells(2, 12).Value
p_increase = Cells(2, 11).Value

For q = 2 To 3001

'Greatest total volume
    If Cells(q, 12).Value > p_change Then
    p_change = Cells(q, 12).Value
    Cells(4, 16).Value = Cells(q, 9).Value
    
    
    Else
    p_change = p_change
    Cells(4, 17).Value = p_change
    End If

'Greatest % increase
    If Cells(q, 11).Value > p_increase Then
    p_increase = Cells(q, 11).Value
    Cells(2, 16).Value = Cells(q, 9).Value
    
    Else
    p_increase = p_increase
    Cells(2, 17).Value = p_increase
    End If

Next q




For q = 2 To 3001
'Greatest % decrease
    If Cells(q, 11).Value < p_increase Then
    p_increase = Cells(q, 11).Value
    Cells(3, 16).Value = Cells(q, 9).Value
    
    Else
    p_increase = p_increase
    Cells(3, 17).Value = p_increase
    End If

Next q




End Sub




    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    






