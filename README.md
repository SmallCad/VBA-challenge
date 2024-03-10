# VBA-challenge

## Using VBA scripting to analyze generated stock market data
In this challenge, I am trying to prepare a VBA script that will analyze a stock's performace after a year; given its Ticker, Dates it fluctuated, the Opening, High, Low, and Closing values of said stock as well as, the Stock volume on the given date.  I would have successfully completed this challenge when my code is able to run accross multiple spreadsheets and provide the analysis of all stocks provided.

## Instructions for users.
Firstly, Clone repository.
In the folder VBA-challenge, open the excel document:
---> Copy of Multiple_year_stock_data
The data is presented such that the first 7 coloums are the Ticker, the Dates it fluctuated, the Opening value, the High, the Low, and the Closing values of the stock. The final coloumn is the total Stock volume on the given date. Here's an example:

ticker      date	   open 	high	low	    close	    vol   'setting up headers is a challenge here, but these are excel headers for numbers below

AAB	       20180102   24.44	    24.56	24.44	24.47	    261879

By pressing the provided "Stock_Analysis" button in the 2018 worksheet, the code should provide a breakdown of every Ticker's performance at the end of the year as well as, which Tickers had the Greatest % increase in value, Greatest % decrease value, and Greatest total volume at the end of the year - across all worksheets.

## Whats under the hood?
Using a for loop, my code sets out to provide the difference between a stock's opening and closing value as well as, provide the percentage change in said stock. It does this by storing a ticker's opening value in a variable. Then, the code loops until it finds its the last interation of a Tickers value i.e. the closing value and it will calculate the difference in value and the percentage changes. Along this recursive loop, a variable, vol, sums up all stock volume from the provided dates for the Ticker. Here is sample code:

    'Variable that the ticker's name is stored in
        Dim ticker As String

    'Variable holding the opening value of the stock value at the beginning of the year
        Dim opener As Double   

    'variable holding the total sum of the volume of stocks for a ticker
        Dim vol As LongLong
        
    'Variable that tracks the row of the ticker, yearly change, percent change, and total stock volume that will be updated
        Dim tracker As Integer


    For i = 2 To last_Row_in_worksheet
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
           Cells(tracker, 10).Value = ticker
           Cells(tracker, 11).Value = Cells(i, 6).Value - opener
           Cells(tracker, 12).Value = (Cells(i, 6).Value / opener) - 1
           Cells(tracker, 13).Value = vol                       
            
            ticker = Cells(i + 1, 1)
            opener = Cells(i + 1, 3)
            vol = Cells(i + 1, 7)
            tracker = tracker + 1
        
        Else: vol = vol + Cells(i + 1, 7)
        End If
        
    Next i


## Credit
I ran into a few set backs during this challenge. 
Firstly, I wasn't sure if I was allowed to do any formatting through excel's control panel or if it all had to be done by VBA.
--> Xpert Learning Assistant - NEW! was used to generate codes for formatting cells as percentages. the code below was pretty much used as is from the AI:
Dim rng As Range
    Set rng = Range("A1") as percentages
    rng.NumberFormat = "0.00%"

Finally, I wasn't sure how to write a README. so thanks to the following links for inspiring this one.

 https://www.freecodecamp.org/news/how-to-write-a-good-readme-file/ -How to Write a Good README File for Your GitHub Project by Hillary Nyakundi

 https://www.youtube.com/watch?v=E6NO0rgFub4 - How To Write a USEFUL README On Github by Ask Cloud Architech
