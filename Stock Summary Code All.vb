vba
'Create 4 columns in all sheets
'loop through each row until a new ticker
'place ticker name in 2,9
'how to do each year????
'pull the first open price and the last close price in the year
'subtract open and close for yearly change
'place value in cell 2,10
'divide close by open format to percent
'place value in cell 2,11
'sum of volume for each day within the year
'place value in cell 2,12
'Repeat for each ticker and move the cell placement by 1 each time.
'Make it work on all sheets in the workbook

Sub Stock_Summary()

'header labels
Dim ws As Worksheet

 For Each ws In Worksheets
    ws.Cells(1, 9).Value = "Ticker"
     ws.Cells(1, 9).Font.Bold = True
    ws.Cells(1, 10).Value = "Yearly Change"
     ws.Cells(1, 10).Font.Bold = True
    ws.Cells(1, 11).Value = "Percent Change"
     ws.Cells(1, 11).Font.Bold = True
    ws.Cells(1, 12).Value = "Total Volume"
     ws.Cells(1, 12).Font.Bold = True

'Declarations
 Dim i As Long
 Dim LastRow As Long
 Dim Year As Date
 Dim YearChange As Double
 Dim OpenPrice As Double
 Dim ClosePrice As Double
 Dim PercentChange As Double
 Dim TotalVolume As Double
 Dim Repeat As Integer
    Repeat = 0
 Dim Offset As Integer
    Offset = 2
    
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'row loops
 For i = 2 To LastRow
  If ws.Cells(i, 1) = ws.Cells(i, 1).Offset(1, 0) Then
   Repeat = Repeat + 1
   TotalVolume = TotalVolume + ws.Cells(i, 7)
   If Repeat = 1 Then
   OpenPrice = ws.Cells(i, 3)
   Else
   End If
   
   Else
   Runtot = Runtot + ws.Cells(i, 1)
   ws.Cells(Offset, 9) = ws.Cells(i, 1)
   ws.Cells(Offset, 12) = TotalVolume
   ClosePrice = ws.Cells(i, 6)
   
   If OpenPrice <> 0 Then
    PercentChange = ((ClosePrice - OpenPrice) / OpenPrice)
    YearChange = ClosePrice - OpenPrice
   Else
    PercentChange = 0
    YearChange = 0
   End If
    
    ws.Cells(Offset, 11) = PercentChange
    ws.Cells(Offset, 11).NumberFormat = "0.00%"
    ws.Cells(Offset, 10) = YearChange
   
   If ws.Cells(Offset, 10).Value > 0 Then
    ws.Cells(Offset, 10).Interior.ColorIndex = 4
   Else
    ws.Cells(Offset, 10).Interior.ColorIndex = 3
   End If
   
   TotalVolume = 0
   Offset = Offset + 1
   Repeat = 0
   
   End If
   
   Next i
   
   Offset = 2
   
Next ws

End Sub

'Autofit columns

Sub AutoFit()

Dim ws As Worksheet

 For Each ws In Worksheets
 ws.Select
     Cells.EntireColumn.AutoFit
    
Next ws
        
End Sub

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

