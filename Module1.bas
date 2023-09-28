Attribute VB_Name = "Module1"
Sub Ticker()

Dim Ticker As String

Dim Yearly_Change As Double
Dim Yearly_Open As Double
Dim Yearly_Close As Double

Dim Percentage_Change As Double

Dim Total_Stock_Volume As Double

Dim Ticker_Summary As Long

Dim G_Increase As Double
Dim G_Decrease As Double
Dim G_Total_Volume As Double
Dim G_Increase_Ticker As String
Dim G_Decrease_Ticker As String
Dim G_Total_Volume_Ticker As String


For Each ws In Worksheets

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

Ticker_Summary = 2
Total_Stock_Volume = 0
Yearly_Open = ws.Range("C2").Value
G_Increase = 0
G_Decrease = 0
G_Total_Volume = 0

lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
For i = 2 To lastRow
Total_Stock_Volume = Total_Stock_Volume + ws.Range("G" & i).Value
  If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    Ticker = ws.Cells(i, 1).Value
    ws.Range("I" & Ticker_Summary).Value = Ticker
    
    Yearly_Close = ws.Cells(i, 6).Value
    
    Yearly_Change = (Yearly_Close - Yearly_Open)
    If Yearly_Open = 0 Then
    Percentage_Change = 0
    Else
    Percentage_Change = Yearly_Change / Yearly_Open
    End If
    ws.Cells(Ticker_Summary, 10).Value = Yearly_Change
    ws.Cells(Ticker_Summary, 11).Value = Percentage_Change
    Yearly_Open = ws.Cells(i + 1, 3).Value
    ws.Cells(Ticker_Summary, 10).NumberFormat = "0.00"
    ws.Cells(Ticker_Summary, 11).NumberFormat = "0.00%"
    
 
  
    ws.Cells(Ticker_Summary, 12).Value = Total_Stock_Volume
    If Percentage_Change > G_Increase Then
        G_Increase = Percentage_Change
        G_Increase_Ticker = Ticker
        ws.Range("Q2").NumberFormat = "0.00%"
    
    End If
    
   

    If ws.Range("J" & Ticker_Summary) > 0 Then
        ws.Range("J" & Ticker_Summary).Interior.ColorIndex = 4
        ElseIf ws.Range("J" & Ticker_Summary) < 0 Then
        ws.Range("J" & Ticker_Summary).Interior.ColorIndex = 3
        Else: ws.Range("J" & Ticker_Summary).Interior.ColorIndex = 0
        
    

    
    End If
    
    If Percentage_Change < G_Decrease Then
        G_Decrease = Percentage_Change
        G_Decrease_Ticker = Ticker
        ws.Range("Q3").NumberFormat = "0.00%"
    End If
    
    If Total_Stock_Volume > G_Total_Volume Then
        G_Total_Volume = Total_Stock_Volume
        G_Total_Volume_Ticker = Ticker
    End If
    
  Total_Stock_Volume = 0
        
  Ticker_Summary = Ticker_Summary + 1
  End If
  
  Next i
  ws.Range("P2").Value = G_Increase_Ticker
  ws.Range("P3").Value = G_Decrease_Ticker
  ws.Range("P4").Value = G_Total_Volume_Ticker
  ws.Range("Q2").Value = G_Increase
  ws.Range("Q3").Value = G_Decrease
  ws.Range("Q4").Value = G_Total_Volume
  Next ws
  
End Sub

