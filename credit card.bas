Attribute VB_Name = "Module1"
Sub CardChecker():

' Range("A1:A101").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("G2"), Unique:=True


Dim Col As Integer
Dim Total As Integer
Dim cc As Integer

cc = 2
Col = 1
Total = 0

    lastrow = Cells(Rows.Count, Col).End(xlUp).Row
    
    For r = 2 To lastrow
    
    Total = Total + Cells(r, 3).Value
    
        If Cells(r + 1, Col).Value <> Cells(r, Col).Value Then
        
            Cells(cc, 7).Value = Cells(r, Col).Value
            Cells(cc, 8).Value = Total
            
            Cells(cc, 7).Interior.ColorIndex = 6
            Cells(cc, 8).Interior.ColorIndex = 6
            
            
            Total = 0
            cc = cc + 1
        End If
     Next r
     


End Sub
