Attribute VB_Name = "Module1"
Sub ChessBoard():
 
 Range("A1:H8").RowHeight = 50
 Range("A1:H8").ColumnWidth = 12
 
 For rownum = 1 To 8
 
    For col = 1 To 8
        
        If col Mod 2 = 1 And rownum Mod 2 = 1 Then
            
            Cells(rownum, col).Interior.ColorIndex = 3
            
        ElseIf col Mod 2 = 1 And rownum Mod 2 = 0 Then
        
            Cells(rownum, col).Interior.ColorIndex = 1
            
        ElseIf col Mod 2 = 0 And rownum Mod 2 = 1 Then
        
            Cells(rownum, col).Interior.ColorIndex = 1
        
        ElseIf col Mod 2 = 0 And rownum Mod 2 = 0 Then
            
            Cells(rownum, col).Interior.ColorIndex = 3

        End If
        
    Next col
    
 Next rownum
 
End Sub
