Attribute VB_Name = "Module1"

Sub stock_analyze()

'looping through worksheets
     
     Dim ws As Worksheet
    
     For Each ws In ThisWorkbook.Worksheets
     
    'Variable ticker's name is stored in
        Dim ticker As String
    
    'Variable holding the opening value of the stock value at the beginning of the year
        Dim opener As Double
        
    'variable holding the total sum of the volume of stocks for a ticker
        Dim vol As LongLong
        
    'Variable that tracks the row of the ticker, yearly change, percent change, and total stock volume that is being updated
        Dim tracker As Integer

    'Variable for the last row in column1 for a given worksheet
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    'initializing ticker as the first ticker in the data
        ticker = ws.Cells(2, 1).Value
        
    'initializing opener as the first open stock value in the data
        opener = ws.Cells(2, 3).Value
        
    'initializing vol as the first volume of stocks in the data
        vol = ws.Cells(2, 7).Value
        
    'initializing tracker as the first row to be updated
        tracker = 2
        
        
    'Creating the headers
        ws.Cells(1, 10).Value = "Tickers"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Total stock Volume"
        ws.Cells(1, 17).Value = "Tickers"
        ws.Cells(1, 18).Value = "Value"
        ws.Cells(2, 16).Value = "Greatest % increase"
        ws.Cells(3, 16).Value = "Greatest % decrease"
        ws.Cells(4, 16).Value = "Greatest total volume"
        
        
    'Obtaining tickers, yearly change, percent change, and total stock volume on the current sheets
    
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ws.Cells(tracker, 10).Value = ticker
            ws.Cells(tracker, 11).Value = ws.Cells(i, 6).Value - opener
            ws.Cells(tracker, 12).Value = (ws.Cells(i, 6).Value / opener) - 1
            ws.Cells(tracker, 13).Value = vol
            
            'Formatting cells. positive percent changes are green. negative and 0 are red.
            If ws.Cells(tracker, 11).Value > 0 Then
                ws.Cells(tracker, 11).Interior.ColorIndex = 4 ' Green
                ws.Cells(tracker, 12).Interior.ColorIndex = 4 ' Green
                
            Else: ws.Cells(tracker, 11).Interior.ColorIndex = 3 ' Red
                ws.Cells(tracker, 12).Interior.ColorIndex = 3 ' Red
            
            End If
                        
            
            ticker = ws.Cells(i + 1, 1)
            opener = ws.Cells(i + 1, 3)
            vol = ws.Cells(i + 1, 7)
            tracker = tracker + 1
        
        Else: vol = vol + Cells(i + 1, 7)
        End If
        
    Next i
 '----------------------------------------------------------------------------------------------------------------------------------------------
            
    'Obtaining the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
    
    'declaring variables to store the "Greatest % increase", "Greatest % decrease", and "Greatest total volume" values
        Dim gpi As Double
        Dim gpd As Double
        Dim gtv As LongLong
    
    'Assumes no two tickers have the same value in any of the three variables.
        Dim Tick1 As String
        Dim Tick2 As String
        Dim Tick3 As String
        
    'Variable representing the last row in column13, "Greatest total volume"
        lastrow13 = Cells(Rows.Count, 13).End(xlUp).Row
    
    'initializing variables as the first data entry for each variable.
        gpi = ws.Cells(2, 12).Value
        gpd = ws.Cells(2, 12).Value
        gtv = ws.Cells(2, 13).Value
        Tick1 = ws.Cells(2, 10).Value
        Tick2 = ws.Cells(2, 10).Value
        Tick3 = ws.Cells(2, 10).Value
        
        
    'finding the gpi, gpd and their tickers
        
        For i = 3 To lastrow13
            If gpi < ws.Cells(i, 12).Value Then
                gpi = ws.Cells(i, 12).Value
                Tick1 = ws.Cells(i, 10)
        
            ElseIf gpd > ws.Cells(i, 12).Value Then
                gpd = ws.Cells(i, 12).Value
                Tick2 = ws.Cells(i, 10)
                    
            End If
            
        Next i
        
        
    'Finding the gtv and its ticker
        
        For j = 3 To lastrow13
            If gtv < ws.Cells(j, 13).Value Then
                gtv = ws.Cells(j, 13).Value
                Tick3 = ws.Cells(j, 10)
            
            End If
            
        Next j
 '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    'display the "Greatest % increase", "Greatest % decrease", and "Greatest total volume" values
        'along with their tickers in the worksheet
    
    
    ws.Range("Q2").Value = Tick1
    ws.Range("Q3").Value = Tick2
    ws.Range("Q4").Value = Tick3
    ws.Range("R2").Value = gpi
    ws.Range("R3").Value = gpd
    ws.Range("R4").Value = gtv
    
    
    
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''Formating data

    'Autofit to display data
    ws.Columns("A:R").AutoFit

    'Declaring arrays to be formatted as percentages for readability
    Dim rng1 As Range
    Dim rng2 As Range
    Set rng1 = ws.Range("L:L")
    Set rng2 = ws.Range("R2:R3")
    
    ' Format the range as percentages
    rng1.NumberFormat = "0.00%"
    rng2.NumberFormat = "0.00%"

    
 Next ws

End Sub

