Attribute VB_Name = "Module1"
Sub Stock_Data()

'Define all variables
Dim ws As Worksheet
Dim Ticker As String
Dim Summary_Table_Row As Integer
Dim Volume As Double
Volume = 0
Dim openprice As Double
Dim closeprice As Double
Dim yearchange As Double

'loop through all worksheets
For Each ws In ThisWorkbook.Worksheets

    'set header values
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'define last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Summary_Table_Row = 2
    
    'nested loop for table
    For i = 2 To LastRow
        
        'conditional
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        'define ticker
        Ticker = Cells(i, 1).Value
        Range("I" & Summary_Table_Row).Value = Ticker
        
        'fill volume
        Volume = Volume + Cells(i, 7).Value
        Range("L" & Summary_Table_Row).Value = Volume
        
        'define columns
        openprice = Cells(i, 3).Value
        closeprice = Cells(i, 6).Value
        
        
        yearchange = closeprice - openprice
        percentagechange = (closeprice - openprice) / closeprice
        
        Range("J" & Summary_Table_Row).Value = yearchange
        Range("K" & Summary_Table_Row).Value = percentagechange
        
        Summary_Table_Row = Summary_Table_Row + 1
        
        'reset volume
        Volume = 0
        
        
        Else
            Volume = Volume + Cells(i, 7).Value
        
    End If
    
    Next i
    
'change column format for percentagechange
Columns("K").NumberFormat = "0.00%"
    
Next ws
    
End Sub
