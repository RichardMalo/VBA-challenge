Sub Stocks():
                
    'Declare all variables used in Sub
    Dim ws As Worksheet
    Dim Ticker_Row As Double
    Dim Total_Stock_Vol As Double
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Next_Ticker As String
    Dim Rows_Count As Double
    Dim RowK_Count As Double
    Dim Greatest_Ticker As String
    Dim Greatest_Percent As Double
    Dim Smallest_Ticker As String
    Dim Smallest_Percent As Double
    Dim Biggest_Volume_Ticker As String
    Dim Biggest_Volume As Double
    
    For Each ws In ThisWorkbook.Worksheets
                        
        With ws
        
        'Count Number of Rows
        Rows_Count = .Cells(Rows.Count, 1).End(xlUp).Row

        'Print out Name Ranges to Excel
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        
        Open_Price = ws.Range("C2")
        Next_Ticker = ws.Range("A2")
           
        'Begin Loop for Totals
        Ticker_Row = 2
        Total_Stock_Vol = 0
           
        'Place First Ticker from Cell(2,1) to Cell (9,1)
        ws.Cells(Ticker_Row, 9).Value = Next_Ticker
        
        For i = 2 To (Rows_Count)
            'Keep Total of Volume
            Total_Stock_Vol = Total_Stock_Vol + ws.Cells(i, 7).Value
            'Check if Stock Symbol is still the same
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ws.Cells(Ticker_Row, 12) = Total_Stock_Vol
            'Then Reset Total_Stock_Volume for the next ticker
            Total_Stock_Vol = 0
            'Display Next Stock_Ticker
            Next_Ticker = ws.Cells(i, 1).Value
            'Grab the next Ticker into memory
            ws.Range("I" & Ticker_Row) = Next_Ticker
            'Record Close_Price Value
            Close_Price = ws.Cells(i, 6).Value
            'Display Differnce between Year Open and Year Close to table
            Yearly_Change = Close_Price - Open_Price
            'Record the Yearly Change into memory
            ws.Cells(Ticker_Row, 10).Value = Yearly_Change
            'Display Percent_Change to table
            Percent_Change = Yearly_Change / Open_Price
            .Cells(Ticker_Row, 11).Value = Percent_Change
            'Add 1 to Ticker_Row to Move to next ticker
            Ticker_Row = Ticker_Row + 1
            'Record Next ticker Open_Price
            Open_Price = ws.Cells(i + 1, 3)
        
            Else
            'If no Ticker change move to the next row and do nothing.
            End If
        
        Next i
               
        'Count Number of Rows in K Column
        RowK_Count = ws.Cells(Rows.Count, 11).End(xlUp).Row
        'Change Cell to 2 decimal points for %
        ws.Range("K2:K" & RowK_Count).NumberFormat = "0.00%"
    
        'Greatest % Increase
        'Greatest Total Volume
        
        Greatest_Percent = 0.000000001
        Smallest_Percent = 100000
        
        For p = 2 To (RowK_Count)
            If (ws.Cells(p, 11).Value > Greatest_Percent) Then
            Greatest_Ticker = ws.Cells(p, 9).Value
            Greatest_Percent = ws.Cells(p, 11).Value
            End If
        Next p
    
        'Print Result of Greatest %
        ws.Range("O2") = Greatest_Ticker
        ws.Range("P2") = Greatest_Percent
        ws.Range("P2").NumberFormat = "0.00%"

        'Greatest % Decrease.
        For q = 2 To (RowK_Count)
            If ws.Cells(q, 11).Value < Smallest_Percent Then
            Smallest_Ticker = ws.Cells(q, 9).Value
            Smallest_Percent = ws.Cells(q, 11).Value
            End If
        Next q
        'Print Result of Greatest % Decrease.
        ws.Range("O3") = Smallest_Ticker
        ws.Range("P3") = Smallest_Percent
        ws.Range("P3").NumberFormat = "0.00%"

        'Greatest Total Volume Display.
        For x = 2 To (RowK_Count)
            If ws.Cells(x, 12).Value > Biggest_Volume Then
            Biggest_Volume_Ticker = ws.Cells(x, 9).Value
            Biggest_Volume = ws.Cells(x, 12).Value
            End If
        Next x

        'Print Result of Greatest Total Volume.
        ws.Range("O4") = Biggest_Volume_Ticker
        ws.Range("P4") = Biggest_Volume
        ws.Range("P4").NumberFormat = "0000000000000"

        'Autofit Columns
        ws.Columns("I:P").AutoFit
           
        'Change Colors to either Red, Green or Yellow
        For j = 2 To (RowK_Count)
            If ws.Cells(j, 10) = 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 6
            ElseIf ws.Cells(j, 10) < 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 3
            ElseIf .Cells(j, 10) > 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 4
            End If
        Next j
        
        'Reset Variables for next Worksheet
        Total_Stock_Vol = 0
        Ticker_Row = 2
        Open_Price = 0
        Close_Price = 0
        Yearly_Change = 0
        Percent_Change = 0
        Smallest_Percent = 10000000
        Greatest_Percent = 0.000000001
        Biggest_Volume = 0
        Rows_Count = 0
        RowK_Count = 0

        End With
        
    Next ws
        
End Sub

