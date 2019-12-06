Attribute VB_Name = "Module1"
Sub stock_market_data():

For Each ws In Worksheets
    
    ' Create Headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Set Variables
    Dim Ticker As String
    
    Dim Stock_Volume As LongLong
    Stock_Volume = 0
    
    Dim Table_Row As Long
    Table_Row = 2
    
    Dim WorksheetName As String
    WorksheetName = ws.Name
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim RowNumber As Long
    RowNumber = 2
    
    ' Start loop to generate info
      For i = 2 To LastRow
        
        Dim Stock_Open As Double
        Stock_Open = ws.Range("C" & RowNumber).Value
        
        Dim Stock_Close As Double
        Stock_Close = ws.Cells(i, 6).Value
        
        Dim Yearly_Change As Double
        
        Dim Percent_Change As Double
        
        ' Generate stock data each time a new Ticker symbol is reached.
          If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
              Yearly_Change = Stock_Close - Stock_Open
              
              If Stock_Open > 0 Then
                Percent_Change = Yearly_Change / Stock_Open
              
              End If
              
              Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
              
              ' Print Ticker symbol
              ws.Range("I" & Table_Row).Value = ws.Cells(i, 1).Value
              
              ' Print Yearly Change & format cells based on values
              ws.Range("J" & Table_Row).Value = Yearly_Change
                  
                  If ws.Range("J" & Table_Row).Value > 0 Then
                  ws.Range("J" & Table_Row).Interior.ColorIndex = 4
                  
                  ElseIf ws.Range("J" & Table_Row).Value < 0 Then
                  ws.Range("J" & Table_Row).Interior.ColorIndex = 3
                  
                  End If
                  
              ' Print Percent Change
              ws.Range("K" & Table_Row).Value = Percent_Change
              ws.Range("K2:K" & LastRow).Style = "Percent"
              ws.Range("K2:K" & LastRow).NumberFormat = "0.00%"
    
              ' Print Stock Volume
              ws.Range("L" & Table_Row).Value = Stock_Volume
              
              ' Set conditions for next i
              Table_Row = Table_Row + 1
              Stock_Volume = 0
              RowNumber = RowNumber + ((i + 1) - RowNumber)
              
          Else: Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
                
                
          End If
    
        
    
      Next i
      
      ' Greatest % Increase, Greatest % Decrease & Greatest Total Volume
      ws.Range("P1").Value = "Ticker"
      ws.Range("Q1").Value = "Value"
      ws.Range("O2").Value = "Greatest % Increase"
      ws.Range("O3").Value = "Greatest % Decrease"
      ws.Range("O4").Value = "Greatest Total Volume"
      ws.Range("Q2:Q3").Style = "Percent"
      ws.Range("Q2:Q3").NumberFormat = "0.00%"
      
        ' Max Value (% increase)
        Value_MaxIncrease = WorksheetFunction.Max(ws.Range("K2:K" & Table_Row))
        ws.Range("Q2").Value = Value_MaxIncrease
        
        '' Ticker corresponding to Max Value (% Increase)
        RowNumber_MaxIncrease = WorksheetFunction.Match(Value_MaxIncrease, ws.Range("K2:K" & Table_Row), 0)
        ws.Range("P2").Value = ws.Range("I" & (RowNumber_MaxIncrease + 1))

        ' Min Value (% Decrease)
        Value_MaxDecrease = WorksheetFunction.Min(ws.Range("K2:K" & Table_Row))
        ws.Range("Q3").Value = Value_MaxDecrease
      
        '' Ticker corresponding to Min Value (% Decrease)
        RowNumber_MaxDecrease = WorksheetFunction.Match(Value_MaxDecrease, ws.Range("K2:K" & Table_Row), 0)
        ws.Range("P3").Value = ws.Range("I" & (RowNumber_MaxDecrease + 1))
        
        ' Max Value (Total Volume)
        Value_MaxVolume = WorksheetFunction.Max(ws.Range("L2:L" & Table_Row))
        ws.Range("Q4").Value = Value_MaxVolume
        
        '' Ticker corresponding to Max Value (Total Volume)
        RowNumber_MaxVolume = WorksheetFunction.Match(Value_MaxVolume, ws.Range("L2:L" & Table_Row), 0)
        ws.Range("P4").Value = ws.Range("I" & (RowNumber_MaxVolume + 1))
      
    

Next ws


End Sub

