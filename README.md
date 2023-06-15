# VBA-challenge

#The following code was based on Lesson Plans, VBA Scripting, Activities, Cells and Ranges. The code directly inserts data in specific ranges/cells - Summary table and Bonus table.

##Code

        #insert data via ws.Ranges for summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Percentage"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"
        
        #insert data via ws.Ranges for bonus table
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"


#The following code was based from the Excel Vba Activities Credit Card Checker. The beginning part of the code sets the variables while the second part of the code generates the Total Stock Volume of each ticker/stock into the summary table.

##Code

      #beginning part of the code

        #set variable for ticker
    Dim ticker As String
        
        #set variable for last row, yearly change, percentage change, opening price, closing price, total stock volume
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
        
            #keep track of location in summary table
      Dim summary_table As Integer
      summary_table = 2



          #second part of the code

              #check if still same ticker
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                                
              #set the ticker name
      ticker = ws.Cells(i, 1).Value

              #add total stock volume
      total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
 
              #print stock name in summary table
      ws.Range("K" & summary_table).Value = ticker
                                                
              #print total stock volume in summary table
      ws.Range("N" & summary_table).Value = total_stock_volume


#The following code snippet and guidance were provided by AskBCs LA, Saad. The provided code helped generate the yearly change of each ticker. It stores the opening price at the beginning of each ticker and the closing price of each ticker at the end of the year. Then it calculates the yearly change for each ticker. Yearly change gets converted into percentage change. The IF condition is to format the yearly change into green if it's a positive and green if it's a negative.

##Code

          #first part

              #keep track of location in bonus table
        Dim openpricerow As Double
        openpricerow = 2

          #second part

              #store the opening and closing price for each ticker at the end of the year
      closing_price = ws.Cells(i, 6).Value
      opening_price = ws.Cells(openpricerow, 3).Value
                                                                            
              #calculate the yearly change
      yearly_change = closing_price - opening_price
                                                            
              #calculate the percentage change
      percentage_change = (closing_price - opening_price) / opening_price
                                
              #print yearly change in summary table
      ws.Range("L" & summary_table).Value = yearly_change
                                                
               #print percentage change in summary table
      ws.Range("M" & summary_table).Value = percentage_change
                                                
              #convert yearly change in summary table
      ws.Range("L" & summary_table).NumberFormat = "0.00"
                                                
 	            #convert percentage change in summary table
      ws.Range("M" & summary_table).NumberFormat = "0.00%"
                                                
                                                
              #change colour
      If ws.Range("L" & summary_table).Value > 0 Then
              ws.Range("L" & summary_table).Interior.ColorIndex = 4
                                                 
      Else
              ws.Range("L" & summary_table).Interior.ColorIndex = 3
                                   
      End If



#The following code is to format the range in Percentage came from Stack overflow. This source provided by AskBCs LA, Saad.

##Code

                #convert yearly change in summary table
      ws.Range("L" & summary_table).NumberFormat = "0.00"
                                                
 	              #convert percentage change in summary table
      ws.Range("M" & summary_table).NumberFormat = "0.00%"


#The following code snippet and guidance provided by AskBCs LA Mohamed. The provided code generates the Greatest % Increase and Decrease, as well as Greatest Total Volume. The code search for the Max and Min number from the entire column of percentage change and stock volume. When it finds the specific values, it will search the column to match the value to a specific ticker. 

##Code

                  #get the max and min and place them in a separate part in the worksheet
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & lastrow)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & lastrow)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & lastrow))

                  #match the max and min values from the range
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastrow)), ws.Range("L2:L" & lastrow), 0)

                  #find the ticker that match the greatest % of increase and decrease, and volume
        ws.Range("P2") = Cells(increase_number + 1, 9)
        ws.Range("P3") = Cells(decrease_number + 1, 9)
        ws.Range("P4") = Cells(volume_number + 1, 9)
![image](https://github.com/janceldawn/VBA-challenge/assets/134527987/14399cfb-1c55-4a03-8fd9-cf488feca1c9)
