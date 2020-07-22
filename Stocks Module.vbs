Attribute VB_Name = "Module1"
Sub stocks()
Dim ws As Worksheet

For Each ws In Worksheets

'set up column headers. Adjusted to add color and autofit contents

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 9).Columns.AutoFit
ws.Cells(1, 9).Interior.ColorIndex = 8
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 10).Columns.AutoFit
ws.Cells(1, 10).Interior.ColorIndex = 8
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 11).Columns.AutoFit
ws.Cells(1, 11).Interior.ColorIndex = 8
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 12).Columns.AutoFit
ws.Cells(1, 12).Interior.ColorIndex = 8

                                      
'identify items we'll need to store for script

Dim closestock As Double
Dim ticker As String
Dim TotalVol As Double
Dim outputrow As Long
outputrow = 2
Dim year_Change As Double
Dim PercentChange As Double

Dim LastRow As Long    ',1 means first column. In column count 1 is for row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'use the For Loop to go through all items for each ticker. Rather than trying to do the whole sheet then go back.

For i = 2 To LastRow

    'Make sure these If statements are all tabbed over correctly to be cought in the first i of the For Loop
    
    If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
    openstock = ws.Cells(i, 3).Value
    
   'this will identify "Ticker A" as the first ticker to deal with. then move to AA and continue to the end.
    
    End If

'Calculate the total volume of stock by adding the value in the first row's volume to each
'subsequent row for that first ticker.

    TotalVol = TotalVol + ws.Cells(i, 7).Value
        
    If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
            'Ticker
            ws.Cells(outputrow, 9).Value = ws.Cells(i, 1).Value
            'Volume
            ws.Cells(outputrow, 12).Value = TotalVol
            'closing stock
            closestock = ws.Cells(i, 6).Value
            'calculate change
            year_Change = closestock - openstock
            ws.Cells(outputrow, 10).Value = year_Change
      
  
            If year_Change >= 0 Then
            ws.Cells(outputrow, 10).Interior.ColorIndex = 4
            Else
            ws.Cells(outputrow, 10).Interior.ColorIndex = 3
    
            End If

'Now that volume for Ticker A is done calculate the change in that volume between first record and last.
'had to go back and add If's to handle zeros in the sheet as they won't divide and causes stack overflow.
     
      
        If openstock = 0 And closestock = 0 Then
            PercentChange = 0
            ws.Cells(outputrow, 11).Value = PercentChange
            ws.Cells(outputrow, 11).NumberFormat = "0.00%"
    
        ElseIf openstock = 0 Then
            Dim NoPercentChange As String
            NoPercentChange = "New"
            ws.Cells(outputrow, 11).Value = NoPercentChange

        Else
            PercentChange = closestock / openstock
            ws.Cells(outputrow, 11).Value = PercentChange
            ws.Cells(outputrow, 11).NumberFormat = "0.00%"
        End If

'make sure to advance output row down by one from 2 to 3 for next ticker

        outputrow = outputrow + 1

'Had to reset counts are it kept including previous ticker info.

        TotalVol = 0
        openstock = 0
        closestock = 0
        year_Change = 0
        PercentChange = 0
 
 'end original If
     End If
  
   
    
Next i

'set up sells for Greatest % increase/Decrease/and volume

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(2, 15).Columns.AutoFit
ws.Cells(2, 15).Interior.ColorIndex = 8

ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(3, 15).Columns.AutoFit
ws.Cells(3, 15).Interior.ColorIndex = 8

ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(4, 15).Columns.AutoFit
ws.Cells(4, 15).Interior.ColorIndex = 8


ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 16).Columns.AutoFit
ws.Cells(1, 16).Interior.ColorIndex = 8


ws.Cells(1, 17).Value = "Value"
ws.Cells(1, 17).Interior.ColorIndex = 8

LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

'set up for loop for next section to look for greatest and least in each column

Dim BestStock As String
Dim BestPercent As Double

BestPercent = ws.Cells(2, 11).Value

Dim WorstStock As String
Dim WorstPercent As Double

WorstPercent = ws.Cells(2, 11).Value

Dim Bestvolume As String
Dim Bestvolumetotal As Double

Bestvolumetotal = ws.Cells(2, 12).Value



For J = 2 To LastRow
    If ws.Cells(J, 11).Value > BestPercent Then
        BestPercent = ws.Cells(J, 11).Value
        BestStock = ws.Cells(J, 9).Value
    End If
    
    If ws.Cells(J, 11).Value < WorstPercent Then
        WorstPercent = ws.Cells(J, 11).Value
        WorstStock = ws.Cells(J, 9).Value
            
    End If
    
    If ws.Cells(J, 12).Value > Bestvolumetotal Then
        Bestvolume = ws.Cells(J, 9).Value
        Bestvolumetotal = ws.Cells(J, 12).Value
    End If
    
Next J

ws.Cells(2, 16).Value = BestStock
ws.Cells(2, 17).Value = BestPercent
ws.Cells(2, 17).NumberFormat = "0.00%"


ws.Cells(3, 16).Value = WorstStock
ws.Cells(3, 17).Value = WorstPercent
ws.Cells(3, 17).NumberFormat = "0.00%"

ws.Cells(4, 16).Value = Bestvolume
ws.Cells(4, 17).Value = Bestvolumetotal
ws.Cells(4, 17).Columns.AutoFit



Next ws

End Sub
