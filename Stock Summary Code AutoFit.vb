vba
'Autofit columns

Sub AutoFit()

Dim ws As Worksheet

 For Each ws In Worksheets
 ws.Select
     Cells.EntireColumn.AutoFit
    
Next ws
        
End Sub