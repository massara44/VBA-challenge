Attribute VB_Name = "Module1"
Sub stock_market()

  Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate

        ' --------------------------------------------
        ' Created a Worksheet
        Dim WorksheetName As String

                
        ' Add the word Ticker to Column "I" Header
        ws.Cells(1, 9).value = "Ticker"
        
        ' Add the word Yearly change to Column "J"  Header
        ws.Cells(1, 10).value = "Opening Price"
        
        ' Add the word Yearly change to Column "J"  Header
        ws.Cells(1, 11).value = "Closing Price"
        
        ' Add the word Yearly change to Column "J"  Header
        ws.Cells(1, 12).value = "Yearly Change"
        
        ' Add the word Percentage Change to Column "K" Header
        ws.Cells(1, 13).value = "Percentage Change"
        
        ' Add the word Total Stock Volume to Column "L"  Header
        ws.Cells(1, 14).value = "Total Stock Volume"

        ' Set an initial variable for holding the Ticker symbol
        Dim ticker_symbol As String
        
        ' Set an initial variable for holding the Yearly change
        Dim yearly_change As Double
        
        ' Set an initial variable for opening and closing price
        Dim op As Double
        Dim cp As Double
                        
        ' Set an initial variable for holding the Percentage Change
        Dim percentage_change As Double

        ' Set an initial variable for holding the Total Stock Volume
        Dim total_stock_volume As Double
        total_stock_volume = 0

        ' Keep track of the location for each ticker symbol in the summary table
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
        current_row = 0
        
                
        ' Determine the Last Row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


        ' Loop through all stock symbols
     For i = 2 To lastrow

         ' Check if we are still within the same ticker symbol, if it is not...
         If ws.Cells(i + 1, 1).value <> Cells(i, 1).value Then

         ' Set the ticker symbol
         ticker_symbol = ws.Cells(i, 1).value

         ' Add to the Total Stock Volume
         total_stock_volume = total_stock_volume + ws.Cells(i, 7).value

         ' Print the Ticker Symbol in the Summary Table
          ws.Range("I" & Summary_Table_Row).value = ticker_symbol
          
                If ws.Cells(i + 1, 1).value <> Cells(i, 1).value Then
                ' Set the Opening price
                  op = ws.Cells(i - current_row, 3).value
    
                 ' Print the Ticker Symbol in the Summary Table
                    ws.Range("J" & Summary_Table_Row).value = op
                    
                summary_row = summary_row + 1
                current_row = 0
    
                ' Set the Closing price
                 cp = ws.Cells(i, 6).value
              
                ' Print the Ticker Symbol in the Summary Table
                    ws.Range("K" & Summary_Table_Row).value = cp
                 
                 End If
              
          
          ' Set the yearly change
          yearly_change = cp - op
          
          ' Print the Yearly change in the Summary Table
          ws.Range("L" & Summary_Table_Row).value = yearly_change
                If ws.Range("L" & Summary_Table_Row).value > 0 Then

                  ' Color the Passing grade green
                ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 4

                 ' Check if stocks have negative yearly change
                 Else
                ' Color for negative yearly change
                  ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
            
                If op And cp <> 0 Then
                
                    ' Set the percentage change
                      percentage_change = (cp - op) / op
                
                
                    ' Print the Ticker Symbol in the Summary Table
                      ws.Range("M" & Summary_Table_Row).value = percentage_change
                      ws.Range("M" & Summary_Table_Row).NumberFormat = "0.00%"
                Else
                      ws.Range("M" & Summary_Table_Row).value = 0
                  
                End If
                   
         ' Print the Total Stock Volume to the Summary Table
          ws.Range("N" & Summary_Table_Row).value = total_stock_volume

         ' Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1
      
         ' Reset the Brand Total and current row
            total_stock_volume = 0
            current_row = 0
            
        
              
          ' If the cell immediately following a row is the same ticker...
        Else

            ' Add to the Total Stock Volume
             total_stock_volume = total_stock_volume + Cells(i, 7).value
            
            'Add one to the current row
            current_row = current_row + 1

        End If

      Next i
      
      ' Determine the Last Row of Yearly Change per WS
        YCLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        ' Add the word Greatest % increase to the row
        ws.Cells(2, 16).value = "Greatest % increase"
        
        ' Add the word Greatest % decrease to the row
        ws.Cells(3, 16).value = "Greatest % decrease"
        
        ' Add the word Greatest total volume to the row
        ws.Cells(4, 16).value = "Greatest total volume"
        
        ' Add the word Percentage Change to Column "Q" Header
        ws.Cells(1, 17).value = "Ticker"
        
        ' Add the word Total Stock Volume to Column "R"  Header
        ws.Cells(1, 18).value = "Value"
    
    For v = 2 To YCLastRow
            If Cells(v, 13).value = Application.WorksheetFunction.max(ws.Range("M2:M" & YCLastRow)) Then
                Cells(2, 17).value = Cells(v, 9).value
                Cells(2, 18).value = Cells(v, 13).value
                Cells(2, 18).NumberFormat = "0.00%"
            ElseIf Cells(v, 13).value = Application.WorksheetFunction.min(ws.Range("M2:M" & YCLastRow)) Then
                Cells(3, 17).value = Cells(v, 9).value
                Cells(3, 18).value = Cells(v, 13).value
                Cells(3, 18).NumberFormat = "0.00%"
            ElseIf Cells(v, 14).value = Application.WorksheetFunction.max(ws.Range("N2:N" & YCLastRow)) Then
                Cells(4, 17).value = Cells(v, 9).value
                Cells(4, 18).value = Cells(v, 14).value
            End If
      
      Next v
      
    ' Autofit to display data
    Columns("A:S").AutoFit

    Next ws

End Sub

