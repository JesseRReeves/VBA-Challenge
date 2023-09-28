Note: Natural Chan helped me through a lot of the harder coding here.  So props to him!  To make sure no plagiarism occurs I am adding him as a source.



Sub Ticker()
                                      'Started out by making all my Variables.  Ticker is a string so we can string the different ticker names together within the code.
Dim Ticker As String 
                                      ' These are used to calculate the yearly change, and percentage change, as well as placing the summation of said data in the correct columns.
Dim Yearly_Change As Double  
Dim Yearly_Open As Double
Dim Yearly_Close As Double

Dim Percentage_Change As Double
                                      'Used to place Total amount of stocks in the correct Column.
Dim Total_Stock_Volume As Double
                                      'Used to place the Tickers in the correct Column.
Dim Ticker_Summary As Long
                                      'Used to set beginning integer to 0 so we can make "<" or ">" based on the greatest percentage increase and the greatest percentage decrease.
Dim G_Increase As Double
Dim G_Decrease As Double
Dim G_Total_Volume As Double          'Used to set beginning integer to 0 so we could use ">" to find the Greatest Total Volume.
Dim G_Increase_Ticker As String          'Used to find the corresponding Ticker to the greatest increase, decrease, and total volume.
Dim G_Decrease_Ticker As String
Dim G_Total_Volume_Ticker As String


For Each ws In Worksheets                          'Connected all worksheets together.

ws.Range("I1").Value = "Ticker"                  'Wrote out "ticker" on spreadsheets
ws.Range("J1").Value = "Yearly Change"           'Wrote out "Yearly Change" on spreadsheets 
ws.Range("K1").Value = "Percent Change"          'Wrote out "Percent Change" on spreadsheets
ws.Range("L1").Value = "Total Stock Volume"      'Wrote out "Total Stock Volume" on spreadsheets.

ws.Range("O2").Value = "Greatest % Increase"        'Wrote out "Greatest % Increase" on spreadsheets.
ws.Range("O3").Value = "Greatest % Decrease"        'Wrote out "Greatest % Decrease" on spreadsheets.
ws.Range("O4").Value = "Greatest Total Volume"      'Wrote out "Greatest Total Volume" on spreadsheets.
ws.Range("P1").Value = "Ticker"                     'Wrote out "Ticker" on spreadsheets.
ws.Range("Q1").Value = "Value"                      'Wrote out "Value" on spreadsheets.

Ticker_Summary = 2                                  'Ticker summary starts on row 2.
Total_Stock_Volume = 0                              'Needs to start at 0 to add up correctly.
Yearly_Open = ws.Range("C2").Value                  'Where Yearly_Open is found.
G_Increase = 0                                      'Needs to start at 0 to add up correctly.
G_Decrease = 0                                      ' "   "
G_Total_Volume = 0                                  ' "   " 

lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row                          'Set the last row used in the script.        Note: Natural helped me with the lastRow.  Source: Natural
For i = 2 To lastRow                                                       'Set the parameters for i
Total_Stock_Volume = Total_Stock_Volume + ws.Range("G" & i).Value          'Adding the Total stock volumes up.                Note: Natural helped me with this.  Source: Natural
  If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then                     'If each cell going down is not equal to the cell above it, then
    Ticker = ws.Cells(i, 1).Value                                              the Ticker changes to the new cells input.
    ws.Range("I" & Ticker_Summary).Value = Ticker                          'Where this information goes on the spreadsheet(Column "I")
    
    Yearly_Close = ws.Cells(i, 6).Value                              'Where to find the yearly close input.
    
    Yearly_Change = (Yearly_Close - Yearly_Open)                     'Equation to find the Yearly Change.  Also used to find the Percentage Change.          Note: Natural helped me here as well.  Source: Natural
    If Yearly_Open = 0 Then                                          'This If statement is used to prevent any errors that would appear due to the number 0 appearing in the 
    Percentage_Change = 0                                            division equation used to find the Percentage Change.
    Else
    Percentage_Change = Yearly_Change / Yearly_Open                    'Equation to find Percentage_Change.
    End If
    ws.Cells(Ticker_Summary, 10).Value = Yearly_Change                'Where to place the Yearly Change information.
    ws.Cells(Ticker_Summary, 11).Value = Percentage_Change            'Where to place the Percentage Change information. 
    Yearly_Open = ws.Cells(i + 1, 3).Value                            'Where to find the yearly open input.
    ws.Cells(Ticker_Summary, 10).NumberFormat = "0.00"                'Formatted the decimal placement in column 10.
    ws.Cells(Ticker_Summary, 11).NumberFormat = "0.00%"               'Formatted the decimal placement as well as changed the numbers to percentages in column 11.
    
 
  
    ws.Cells(Ticker_Summary, 12).Value = Total_Stock_Volume           'Where to place the Total Stock Volume Summation.          Note: Natural helped me through this process. Source: Natural
    If Percentage_Change > G_Increase Then                            'The equation to find the Greatest percentage increase.
        G_Increase = Percentage_Change                                
        G_Increase_Ticker = Ticker                                    
        ws.Range("Q2").NumberFormat = "0.00%"                         'Changed Number Format to a percentage with two decimals.
    
    End If
    
   

    If ws.Range("J" & Ticker_Summary) > 0 Then                              'Colored the positive Yearly Change summations Green.
        ws.Range("J" & Ticker_Summary).Interior.ColorIndex = 4
        ElseIf ws.Range("J" & Ticker_Summary) < 0 Then                      'Colored the negative Yearly Change summations Red.
        ws.Range("J" & Ticker_Summary).Interior.ColorIndex = 3              
        Else: ws.Range("J" & Ticker_Summary).Interior.ColorIndex = 0        'Colored the Yearly Change summations that = 0 to white.
        
    

    
    End If
    
    If Percentage_Change < G_Decrease Then                                  'The equation to find the greatest percentage decrease.
        G_Decrease = Percentage_Change
        G_Decrease_Ticker = Ticker                                         
        ws.Range("Q3").NumberFormat = "0.00%"                               'Change Number Format to a percentage with two decimals.
    End If
    
    If Total_Stock_Volume > G_Total_Volume Then                             'The equation to find the Greatest Total Volume.
        G_Total_Volume = Total_Stock_Volume
        G_Total_Volume_Ticker = Ticker                                      
    End If
    
  Total_Stock_Volume = 0                                                     'Total_Stock_Volume must start as zero for the equation to work right.
        
  Ticker_Summary = Ticker_Summary + 1                                        'Placing the summation of data down the column.
  End If
  
  Next i
  ws.Range("P2").Value = G_Increase_Ticker                           'Where to place the corresponding Ticker for the G_Increase_Ticker.
  ws.Range("P3").Value = G_Decrease_Ticker                           'Where to place the corresponding Ticker for the G_Decrease_Ticker.
  ws.Range("P4").Value = G_Total_Volume_Ticker                       'Where to place the corresponding Ticker for the G_Total_Volume_Ticker.
  ws.Range("Q2").Value = G_Increase                                  'Where to place the corresponding percentage for th eG_Increase.
  ws.Range("Q3").Value = G_Decrease                                  'Where to place the corresponding percentage for the greatest percentage decrease.
  ws.Range("Q4").Value = G_Total_Volume                              'Where to place the corresponding number for the Greatest Total Volume.
  Next ws
  
End Sub
