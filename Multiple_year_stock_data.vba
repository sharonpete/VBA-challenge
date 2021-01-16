Attribute VB_Name = "Module1"
 Sub Calculate()
    ' funny thing... this initial MsgBox brings the spreadsheet to the fore, here for convenience
    MsgBox ("here we go ...")
      
    For Each ws In Worksheets
    
        Dim WorksheetName As String
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        WorksheetName = ws.Name
        
        MsgBox ("The worksheet name is: " + WorksheetName)
      
    
       ' find the number of rows in this worksheet
       number_rows = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
       
       MsgBox ("number of rows: " + Str(number_rows))
       
       ' initialize the column headers for the summary table
       ws.Cells(1, 9).Value = "Ticker"
       ws.Cells(1, 10).Value = "Yearly Change"
       ws.Cells(1, 11).Value = "Percent Change"
       ws.Cells(1, 12).Value = "Total Stock Volume"
       
       Dim summary_row As Integer
       summary_row = 2
    
       Dim ticker As String
       ticker = ws.Cells(2, 1).Value ' initialize the value with the first row
       ' MsgBox ("first ticker symbol is: " + ticker)
       
       Dim opening_price As Double
       Dim closing_price As Double
       Dim percent_price_change As Double
       Dim volume As LongLong
           
       ' set up variables for bonus work
       Dim summary_table_rows As Long
       Dim greatest_percent_increase As Double
       Dim greatest_percent_decrease As Double
       Dim greatest_total_volume As LongLong
       Dim ticker_increase As String
       Dim ticker_decrease As String
       Dim ticker_volume As String

        
       redColor = 3
       greenColor = 4
       
       volume = 0
       opening_price = ws.Cells(2, 3).Value ' initialize the opening price with the first row
       
       
       For i = 2 To number_rows
                  
           ' Calculate the total stock volume for the year
           volume = volume + ws.Cells(i, 7).Value
           
           
           ' If the next ticker symbol is different, find the closing, calc the percent change
           ' do some stuff
           ' reinitialize the variables
           ' Grab the ticker symbol and closing price from the ...
           If ws.Cells(i + 1, 1).Value <> ticker Then
               'Debug.Print ("ticker symbol: " + ticker)
               
               closing_price = ws.Cells(i, 6).Value
               
                    
               ' Calculate the percent change from yearly open to close price
               yearly_change = closing_price - opening_price
               
               ' Do not divide by 0 (zero)
               If opening_price <> 0 Then
               
                    percent_price_change = (closing_price - opening_price) / opening_price
                    
               End If
               
               
               ' Add to the summary table
               ws.Cells(summary_row, 9).Value = ticker
               ws.Cells(summary_row, 10).Value = yearly_change
               ws.Cells(summary_row, 11).Value = percent_price_change
               ws.Cells(summary_row, 11).NumberFormat = "0.00%"
               ws.Cells(summary_row, 12).Value = volume
               
               ' Add the formatting for positive change (green) and negative change (red)
               If yearly_change < 0 Then
                   ws.Cells(summary_row, 10).Interior.ColorIndex = redColor
               Else
                   ws.Cells(summary_row, 10).Interior.ColorIndex = greenColor
               End If
   
               'grab the new ticker and opening price, reset volume to 0
               ticker = ws.Cells(i + 1, 1).Value
               opening_price = ws.Cells(i + 1, 3).Value
               volume = 0
               summary_row = summary_row + 1
               
           
           End If
            
           
       Next i
       
       ' Bonus section
       ' format bonus summary table
       
       ' number of rows in summary table
       summary_table_rows = ws.Cells(Rows.Count, 9).End(xlUp).Row
       MsgBox ("Rows in summary table: " + Str(summary_table_rows))
       ws.Cells(1, 15).Value = "Ticker"
       ws.Cells(1, 16).Value = "Value"
       
       ' reset variables for each worksheet  (TODO: not sure of variable scope, should ask about this)
       greatest_percent_increase = 0
       greatest_percent_decrease = 0
       greatest_total_volume = 0
       ticker_increase = ""
       ticker_decrease = ""
       ticker_volume = ""
       
       For r = 2 To summary_table_rows
            
            ' find greatest_percent_increase
            If ws.Cells(r, 11).Value > greatest_percent_increase Then
                greatest_percent_increase = ws.Cells(r, 11).Value
                ticker_increase = ws.Cells(r, 9).Value
                
            End If
            
            ' find greatest_percent_decrease
            If ws.Cells(r, 11).Value < greatest_percent_decrease Then
                greatest_percent_decrease = ws.Cells(r, 11).Value
                ticker_decrease = ws.Cells(r, 9).Value
                
            End If
            
            ' find greatest_total_volume
            If ws.Cells(r, 12).Value > greatest_total_volume Then
                greatest_total_volume = ws.Cells(r, 12).Value
                ticker_volume = ws.Cells(r, 9).Value
            End If
       
       Next r
       
       ' populate bonus summary table
       ws.Cells(2, 14).Value = "Greatest Percent Increase"
       ws.Cells(2, 15).Value = ticker_increase
       ws.Cells(2, 16).Value = greatest_percent_increase
       ws.Cells(2, 16).NumberFormat = "0.00%"
       
       ws.Cells(3, 14).Value = "Greatest Percent Decrease"
       ws.Cells(3, 15).Value = ticker_decrease
       ws.Cells(3, 16).Value = greatest_percent_decrease
       ws.Cells(3, 16).NumberFormat = "0.00%"
       
       ws.Cells(4, 14).Value = "Greatest Total Volume"
       ws.Cells(4, 15).Value = ticker_volume
       ws.Cells(4, 16).Value = greatest_total_volume
       
    Next
       
    MsgBox ("Complete...")
    
    ' Q for the Bonus section... is this asking for the statistics overall or just for the worksheet?
    ' Guessing overall
    
 
 
 End Sub

