Attribute VB_Name = "Module1"
Sub HornetNest():

Dim Hornets As String
Dim Bugs As String
Dim Bees As String
Hornets = 0
Bugs = 10
Bees = 5

'Loop through each row

    For rownum = 1 To 6
    
'while in each row, loop through each colum
    
    For col = 1 To 7
    
'if column contains hornets then...

    If Cells(rownum, col).Value = "Hornets" Then
        Hornets = Hornets + 1
                        
        If Bugs <> 0 Then
            Cells(rownum, col).Value = "Bugs"
            Bugs = Bugs - 1
            Cells(2, 9).Value = Bugs
                
        ElseIf Bees <> 0 Then
            Cells(rownum, col).Value = "Bees"
            Bees = Bees - 1
            Cells(2, 10).Value = Bees
        
        End If

    End If
    
  Next col
   
 Next rownum

MsgBox ("we still have hornets")

End Sub
