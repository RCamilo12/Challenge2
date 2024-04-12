Sub Stocks()

Dim openv As Double
Dim closev As Double
Dim ychange As Double
Dim pchange As Double
Dim cRow As Long
Dim result As Long
Dim total As Double
Dim ws As Worksheet
Dim max As Double
Dim min As Double

For Each ws In ThisWorkbook.Worksheets
    With ws
    

    .Range("I1:N1").Value = Array("Ticker", "Close Value", "Open Value", "Yearly Change", "Percentage Change", "Total Stock Volume")

    .Range("P2").Value = "Greatest Increase"
    .Range("P3").Value = "Greatest Decrease"
    .Range("P4").Value = "Greatest Total Stock Volume"
    .Range("Q1").Value = "Ticker"
    .Range("R1").Value = "Value"
    
    cRow = Cells(Rows.Count, 1).End(xlUp).Row 'https://www.wallstreetmojo.com/vba-row-count/
    result = 2
    
    For i = 2 To cRow
    
        
        If .Cells(i + 1, 1).Value <> .Cells(i, 1).Value Then 'Evaluating next row
            .Cells(result, 9).Value = .Cells(i, 1).Value 'Printing tick
            closev = .Cells(i, 6).Value 'Receiving Close value
            .Cells(result, 10).Value = closev 'prints close value
                                                   
        result = result + 1
        End If
        
        If .Cells(i - 1, 1).Value <> .Cells(i, 1).Value Then 'Evaluating previous column
            
            total = 0
            
            openv = .Cells(i, 3).Value 'Receiving Opening value
            .Cells(result, 11).Value = openv 'Printing Opening Value
            
            ychange = .Cells(result, 10).Value - .Cells(result, 11).Value 'Calculating year change
            .Cells(result, 12).Value = ychange 'Print Year Change
            
            pchange = ychange / openv 'Calculates Percentage Change
            .Cells(result, 13).NumberFormat = "0.00%" 'Adds Format to the percentage https://www.statology.org/vba-percentage-format/
            .Cells(result, 13) = pchange 'Prints the percentage
                
                If .Cells(result, 12) > 0 Then 'This If creats the color for the Year Change https://www.excel-easy.com/vba/examples/background-colors.html
                .Cells(result, 12).Interior.ColorIndex = 4
                ElseIf .Cells(result, 12) < 0 Then
                .Cells(result, 12).Interior.ColorIndex = 3
                End If
                                    
            total = total + .Cells(i, 7).Value 'Calculates total
            .Cells(result, 14).Value = total 'Prints total
            
            'Calculates max value in Percentage
            max = Application.WorksheetFunction.max(Range("M1:M" & cRow))
            .Cells(2, 18).NumberFormat = "0.00%"
            .Cells(2, 18).Value = max
            'Calculates min value in Percentage
            min = Application.WorksheetFunction.min(Range("M1:M" & cRow))
            .Cells(3, 18).NumberFormat = "0.00%"
            .Cells(3, 18).Value = min
            'Calculates max value in Total Stock Volume
            max = Application.WorksheetFunction.max(Range("N1:N" & cRow))
            .Cells(4, 18).NumberFormat = "General"
            .Cells(4, 18).Value = max
            
       End If
           
            
            
    Next i
    
    
    End With
Next ws
    
    End Sub
