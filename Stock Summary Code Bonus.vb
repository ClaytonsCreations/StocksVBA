vba

'create labels and headers
'have it look at the summary data and find the greatest increase in PercentChange and copy ticker and percent to new cells
'have it look at the summary data and find the greatest decreas in the PercentChange and copy ticker and percent to new cells
'Have it look at the summary data and find the largest Total Volume and copy ticker and value to new cells


Sub Bonus()
'labels
 Dim ws As Worksheet
    
 For Each ws In Worksheets
  ws.Cells(1, 15).Value = "Ticker"
   ws.Cells(1, 15).Font.Bold = True
  ws.Cells(1, 16).Value = "Value"
   ws.Cells(1, 16).Font.Bold = True
  ws.Cells(2, 14).Value = "Greatest % Increase"
   ws.Cells(2, 14).Font.Bold = True
   ws.Cells(2, 14).Font.Italic = True
  ws.Cells(3, 14).Value = "Greatest % Decrease"
   ws.Cells(3, 14).Font.Bold = True
   ws.Cells(3, 14).Font.Italic = True
  ws.Cells(4, 14).Value = "Greatest Total Volume"
   ws.Cells(4, 14).Font.Bold = True
   ws.Cells(4, 14).Font.Italic = True
   
'Declarations
  Dim i As Long
  Dim PositiveChange As Double
    PositiveChange = 0
  Dim NegativeChange As Double
    NegativeChange = 0
  Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
  Dim LowTick As String
  Dim HiTick As String
  Dim HighVolume As Double
    HighVolume = 0
  Dim VolumeTick As String
  
  
'loop
 For i = 2 To LastRow
 If ws.Cells(i, 11).Value > PositiveChange Then
    PositiveChange = ws.Cells(i, 11).Value
    HiTick = ws.Cells(i, 9).Value
 End If

 If ws.Cells(i, 11).Value < NegativeChange Then
    NegativeChange = ws.Cells(i, 11).Value
    LowTick = ws.Cells(i, 9).Value
 End If
 
 If ws.Cells(i, 12).Value > HighVolume Then
    HighVolume = ws.Cells(i, 12).Value
    VolumeTick = ws.Cells(i, 9).Value
 End If

    ws.Cells(2, 15).Value = HiTick
    ws.Cells(2, 16).Value = PositiveChange
    ws.Cells(2, 16).NumberFormat = "0.00%"
    ws.Cells(3, 15).Value = LowTick
    ws.Cells(3, 16).Value = NegativeChange
    ws.Cells(3, 16).NumberFormat = "0.00%"
    ws.Cells(4, 15).Value = VolumeTick
    ws.Cells(4, 16).Value = HighVolume
    
Next i

Next ws

End Sub

