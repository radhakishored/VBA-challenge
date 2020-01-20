

Sub Stock_Market()

Dim Sum_row As Long
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greatest_Total As Double
Dim Greatest_Increase_Ticker As String
Dim Greatest_Decrease_Ticker As String
Dim Greatest_Total_Ticker As String

    
For Each ws In Worksheets
    ' Create a Variable to Last Row
    Dim WorksheetName As String
    ' Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ' Grabbed the WorksheetName
    WorksheetName = ws.Name

    'Create new columns for summary
    Sum_row = 2
    ws.Cells(1, 9).Value = "Ticker "
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock volume"
     
        'get ticker name and opening price
    Ticker_Name = ws.Cells(2, 1).Value
    Opening_price = ws.Cells(2, 3).Value
    volume = 0
        
        

    For i = 2 To LastRow
       
        ' Check if we are still within the same ticker , if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
         
            volume = volume + ws.Cells(i, 7)
            Closing_price = ws.Cells(i, 6)
            'Write  to summary
            ws.Cells(Sum_row, 9).Value = Ticker_Name
            ws.Cells(Sum_row, 10).Value = Closing_price - Opening_price
            If Closing_price - Opening_price >= 0 Then
                ws.Cells(Sum_row, 10).Interior.ColorIndex = 4 ' Green for Positive change
            Else
                ws.Cells(Sum_row, 10).Interior.ColorIndex = 3 ' Red for negetive change
            End If
            If Opening_price <> 0 Then
                ws.Cells(Sum_row, 11).Value = FormatPercent((Closing_price - Opening_price) / Opening_price)
               
            End If
                
            ws.Cells(Sum_row, 12).Value = volume
            Sum_row = Sum_row + 1
          'Initialize  ticker name and opening price
            Ticker_Name = ws.Cells(i + 1, 1).Value
            Opening_price = ws.Cells(i + 1, 3).Value
            volume = 0
        Else

            ' Add to the Brand Total
            volume = volume + ws.Cells(i, 7).Value

        End If

    Next i
    'Generate greatest values
    LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    Greatest_Increase = Val(ws.Cells(2, 11))
    Greatest_Decrease = Val(ws.Cells(2, 11))
    Greatest_Total = ws.Cells(2, 12)
    For i = 3 To LastRow
    
        'Search for greatest values
        
        If Val(ws.Cells(i, 11)) > Greatest_Increase Then
            Greatest_Increase = Val(ws.Cells(i, 11))
            Greatest_Increase_Ticker = ws.Cells(i, 9)
            
        End If
        
        If Val(ws.Cells(i, 11)) < Greatest_Decrease Then
            Greatest_Decrease = Val(ws.Cells(i, 11))
            Greatest_Decrease_Ticker = ws.Cells(i, 9)
        End If
        
         If Val(ws.Cells(i, 12)) > Greatest_Total Then
            Greatest_Total = ws.Cells(i, 12)
            Greatest_Total_Ticker = ws.Cells(i, 9)
        End If
        
    
    
    Next i
    'Define new columns
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"
    'Greatest Lables
    ws.Cells(2, 15) = "Greatest % Increase"
    ws.Cells(3, 15) = "Greatest %  Decrease"
    ws.Cells(4, 15) = "Greatest Total Valume"
    'Tickers
    ws.Cells(2, 16) = Greatest_Increase_Ticker
    ws.Cells(3, 16) = Greatest_Decrease_Ticker
    ws.Cells(4, 16) = Greatest_Total_Ticker
    'Values
    ws.Cells(2, 17) = FormatPercent(Greatest_Increase)
    ws.Cells(3, 17) = FormatPercent(Greatest_Decrease)
    ws.Cells(4, 17) = Greatest_Total
    
Next ws
End Sub
