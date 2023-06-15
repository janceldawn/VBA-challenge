VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub multiple_stockdata()

            Dim ws As Worksheet
            
            For Each ws In Worksheets
            
            'insert data via ws.Ranges for summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Percentage"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
            'insert data via ws.Ranges for bonus table
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

            'set variable for ticker
        Dim ticker As String
        
            'set variable for last row, yearly change, percentage change, opening price, closing price, total stock volume
        Dim lastrow As Long
        Dim yearly_change As Double
        Dim percentage_change As Double
        Dim opening_price As Double
        Dim closing_price As Double
        Dim total_stock_volume As Double
       
        yearly_change = 0
        percentage_change = 0
        total_stock_volume = 0
        opening_price = 0
        closing_price = 0
        
            'keep track of location in summary table
        Dim summary_table As Integer
        summary_table = 2
        Dim openpricerow As Double
        openpricerow = 2
        
            'determine last row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
                            'loop through all stocks
                    For i = 2 To lastrow
                    
                                                            'check if still same ticker
                                            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                                
                                                                            'set the ticker name
                                                            ticker = ws.Cells(i, 1).Value
                                                
                                                                            'add total stock volume
                                                            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                                                
                                                                            'print ticker/stock name in summary table
                                                            ws.Range("I" & summary_table).Value = ticker
                                                
                                                                            'print total stock volume in summary table
                                                            ws.Range("L" & summary_table).Value = total_stock_volume
                                                
                                                                                            'store the opening and closing price for each ticker at the end of the year
                                                                            closing_price = ws.Cells(i, 6).Value
                                                                            opening_price = ws.Cells(openpricerow, 3).Value
                                                                            
                                                                                            'calculate the yearly change
                                                                            yearly_change = closing_price - opening_price
                                                            
                                                                                            'calculate the percentage change
                                                                            percentage_change = (closing_price - opening_price) / opening_price
                                
                                                                                            'print yearly change in summary table
                                                                            ws.Range("J" & summary_table).Value = yearly_change
                                                
                                                                                            'print percentage change in summary table
                                                                            ws.Range("K" & summary_table).Value = percentage_change
                                                
                                                                                            'convert yearly change in summary table
                                                                            ws.Range("J" & summary_table).NumberFormat = "0.00"
                                                
                                                                                            'convert percentage change in summary table
                                                                            ws.Range("K" & summary_table).NumberFormat = "0.00%"
                                                
                                                
                                                                                                            'change colour
                                                                                    If ws.Range("J" & summary_table).Value > 0 Then
                                                                                                ws.Range("J" & summary_table).Interior.ColorIndex = 4
                                                 
                                                                                    Else
                                                                                                ws.Range("J" & summary_table).Interior.ColorIndex = 3
                                   
                                                                                    End If
                                                
                                                
                                                                            'add one to the summary table
                                                                summary_table = summary_table + 1
                                                
                                                                            'reset total stock volume and opening price
                                                                total_stock_volume = 0
                                                                openpricerow = i + 1
                                                
                                    Else
                                                
                                                                            'add to total stock volume
                                                                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                                                
                                    End If
                                     
                                        
            Next i
            
            
                            'Bonus section
                            'get the max and min and place them in a separate part in the worksheet
                    ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & lastrow)) * 100
                    ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & lastrow)) * 100
                    ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & lastrow))

                            'match the max and min values from the range
                    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
                    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
                    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastrow)), ws.Range("L2:L" & lastrow), 0)

                            'find the ticker that match the greatest % of increase and decrease, and volume
                    ws.Range("P2") = Cells(increase_number + 1, 9)
                    ws.Range("P3") = Cells(decrease_number + 1, 9)
                    ws.Range("P4") = Cells(volume_number + 1, 9)
         
            
            
       Next ws
       

End Sub

