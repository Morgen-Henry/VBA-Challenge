VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_data()
    'Create headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'Define last row in sheet with data
    With ActiveSheet
        last_row = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    
    'Set starting values
    ticker_symbol = Cells(2, 1).Value
    open_price = Cells(2, 3).Value
    close_price = 0
    total_vol = 0
    j = 2
    maxi = 0
    mini = 0
    max_vol = 0
    
    'Loop to (1) add up group tickers (2) find yearly change for these groups (3) find the percent change for these groups (4) find the total stock volume for these groups
    For i = 2 To last_row:
        
        'If this ticker has been seen before
        If Cells(i, 1).Value = ticker_symbol Then
            'reassert the ticker name & insert into sheet
            Cells(j, 9).Value = Cells(i, 1).Value
            'update to include closing price
            close_price = Cells(i, 6).Value
            'update yearly change to be NEW closing price - opening price
            open_close = close_price - open_price
            Cells(j, 10).Value = open_close
            'update percent change
            Cells(j, 11).Value = FormatPercent(open_close / open_price)
            'update total volume to include new data
            total_vol = total_vol + Cells(i, 7).Value
            Cells(j, 12).Value = total_vol
            
        'If this ticker has NOT been seen before
        Else
            'set opening price to the new 1st open value
            open_price = Cells(i, 3).Value
            'update to include closing price
            close_price = Cells(i, 6).Value
            'move j to 1 lower row
            j = j + 1
            'insert new ticker symbol
            ticker_symbol = Cells(i, 1).Value
            Cells(j, 9).Value = ticker_symbol
            'reset total volume
            total_vol = Cells(i, 7).Value
        End If
        
        'conditional formatting for the yearly change column
        If Cells(i, 10).Value > 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        Else
            Cells(i, 10).Interior.ColorIndex = 3
        End If
        
        'find i of max incr/decr values
        If Cells(i, 11).Value > maxi Then
            maxi = Cells(i, 11).Value
            max_tic = Cells(i, 9).Value
        ElseIf Cells(i, 11).Value < mini Then
            mini = Cells(i, 11).Value
            min_tic = Cells(i, 9).Value
        End If
        
        'find max total vol
        If Cells(i, 12).Value > max_vol Then
            max_vol = Cells(i, 12).Value
            vol_tic = Cells(i, 9).Value
        End If
    Next i
    
    'Create Row titles
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    'Create Header
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    
    'find the greaest % increase
    Cells(2, 16).Value = FormatPercent(maxi)
    Cells(2, 15).Value = max_tic
    
    'find the greatest % decrease
    Cells(3, 16).Value = FormatPercent(mini)
    Cells(3, 15).Value = min_tic
    
    'find the greatest total volume
    Cells(4, 16).Value = max_vol
    Cells(4, 15).Value = vol_tic
End Sub









