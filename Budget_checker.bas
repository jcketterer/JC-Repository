Attribute VB_Name = "Module1"
Sub BudgetChecker():

'setting variables

    Dim Budget As Double
        Budget = 100
    
    Dim Price As Double
        Price = 85
        
    Dim Fee As Double
        Fee = 0.15
        
    Dim Total As Double
        Total = Price
      
      
'setting values to variables

    Budget = Range("C3").Value
    Price = Range("F3").Value
    Fee = Range("H3").Value
    
' Setting function for variables and setting total

    Total = Price * (1 + Fee)
                
    Cells(3, 12).Value = Total


' Setting conditonal for message boxes

    If Total > Budget Then
        MsgBox ("over")
        
    ElseIf Total < Budget Then
        MsgBox ("under")
    
    Else
        MsgBox ("right one budget")
        
    End If

' Setting new variable and if to reset to be right on budget

    Dim NewPrice As Double
    NewPrice = Budget / (1 + Range("H3").Value)
    
    If Total > Budget Then
    
    Range("F3").Value = NewPrice
    
    Range("L3").Value = NewPrice * (1 + Range("H3").Value)
    
    End If


End Sub
