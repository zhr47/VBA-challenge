Attribute VB_Name = "Module1"
Sub VBAStockHWZaheer()
   
   ' Set CurrentWs as a worksheet object variable.
    Dim CurrentWs As Worksheet
    Dim summary_table_header As Double
    Dim command_spreadsheet As Double
    
    summary_table_header = False
    command_spreadsheet = True
    
    ' Loop through the worksheets in the workbook.
    For Each CurrentWs In Worksheets
    
        'initial variable for holding the ticker name
        Dim ticker_name As String
        ticker_name = " "
        
        'initial variable for holding the total ticker volume
        Dim total_ticker_volume As Double
        total_ticker_volume = 0
        
        'variables required
        Dim open_price As Double
        open_price = 0
        Dim close_price As Double
        close_price = 0
        Dim yearly_change As Double
        yearly_change = 0
        Dim percent_change As Double
        percent_change = 0
        
        Dim max_ticker_name As String
        max_ticker_name = " "
        Dim min_ticker_name As String
        min_ticker_name = " "
        Dim max_percent As Double
        max_percent = 0
        Dim min_percent As Double
        min_percent = 0
        Dim max_volume_ticker As String
        max_volume_ticker = " "
        Dim max_volume As Double
        max_volume = 0

         
        'tracks of the location for each ticker name
        Dim summary_table_row As Long
        summary_table_row = 2
        
        ' Set initial row count for the current worksheet
        Dim Lastrow As Double
        Lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row
        
        'summary table values
        If summary_table_header Then
            'summary table titles
            CurrentWs.Range("I1").Value = "Ticker"
            CurrentWs.Range("J1").Value = "Yearly Change"
            CurrentWs.Range("K1").Value = "Percent Change"
            CurrentWs.Range("L1").Value = "Total Stock Volume"
            'titles for hard part
            CurrentWs.Range("O2").Value = "Greatest % Increase"
            CurrentWs.Range("O3").Value = "Greatest % Decrease"
            CurrentWs.Range("O4").Value = "Greatest Total Volume"
            CurrentWs.Range("P1").Value = "Ticker"
            CurrentWs.Range("Q1").Value = "Value"
        Else
            summary_table_header = True
        End If
        
        'setting value of open price
        open_price = CurrentWs.Cells(2, 3).Value
        
        'setting i
        Dim i As Double
        For i = 2 To Lastrow
        
            'same ticker?
            If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
            
                'set the ticker name
                ticker_name = CurrentWs.Cells(i, 1).Value
                
                'calculate yearly_change and percent_change
                close_price = CurrentWs.Cells(i, 6).Value
                yearly_change = close_price - open_price
                
                'check Division by 0
                If open_price <> 0 Then
                    percent_change = (yearly_change / open_price) * 100
                Else
                End If
                
                'add to the Ticker name total volume
                total_ticker_volume = total_ticker_volume + CurrentWs.Cells(i, 7).Value
              
                
                'print the Ticker Name
                CurrentWs.Range("I" & summary_table_row).Value = ticker_name
                'print the Yearly Change
                CurrentWs.Range("J" & summary_table_row).Value = yearly_change
                
                'fill "Yearly Change" columns with colors 4=green 3=red
                If (yearly_change > 0) Then
                    CurrentWs.Range("J" & summary_table_row).Interior.ColorIndex = 4
                ElseIf (yearly_change <= 0) Then
                    CurrentWs.Range("J" & summary_table_row).Interior.ColorIndex = 3
                End If
                
                 'columns I &J
                CurrentWs.Range("K" & summary_table_row).Value = (CStr(percent_change) & "%")
                CurrentWs.Range("L" & summary_table_row).Value = total_ticker_volume
                
                'add 1 to the summary table row count
                summary_table_row = summary_table_row + 1
                'reset yearly_change and percent_change holders,
                yearly_change = 0
                'hard part,do this in the beginning of the for loop percent_change = 0
                close_price = 0
                'hold next ticker's open_price
                open_price = CurrentWs.Cells(i + 1, 3).Value
              
                
                'hard part, new Summary table on the right with calcalations
                If (percent_change > max_percent) Then
                    max_percent = percent_change
                    max_ticker_name = ticker_name
                ElseIf (percent_change < min_percent) Then
                    min_percent = percent_change
                    min_ticker_name = ticker_name
                End If
                       
                If (total_ticker_volume > max_volume) Then
                    max_volume = total_ticker_volume
                    max_volume_ticker = ticker_name
                End If
                
                'hard part adjustments to resetting counters
                percent_change = 0
                total_ticker_volume = 0
                
            
            'if same ticker name add to ticker vol, then increase
            Else
                total_ticker_volume = total_ticker_volume + CurrentWs.Cells(i, 7).Value
            End If
      
        Next i

            'hard solution
            If Not command_spreadsheet Then
            
                CurrentWs.Range("Q2").Value = (CStr(max_percent) & "%")
                CurrentWs.Range("Q3").Value = (CStr(min_percent) & "%")
                CurrentWs.Range("P2").Value = max_ticker_name
                CurrentWs.Range("P3").Value = min_ticker_name
                CurrentWs.Range("Q4").Value = max_volume
                CurrentWs.Range("P4").Value = max_volume_ticker
                
            Else
                command_spreadsheet = False
            End If
        
     Next CurrentWs
End Sub

