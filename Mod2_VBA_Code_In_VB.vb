Sub all_sheets()

    'Create a variable to hold the counter
    Dim ws As Worksheet
    
    'Loop though each worksheet
    
    'Run the stock analysis on worksheet
    

End Sub

Sub Analyze_Stock_Data()

'Declare variables
    Dim i As Integer
    Dim Row, Volume, Result As Integer
    Dim Open_Price, End_Price, Quart_Change, Perc_Change As Double
    Dim Ticker As String
    
'Initialize variables
    Row = 2
    Volume = 0
    Result = 2
    Open_Price = 0
    End_Price = 0
    Quart_Change = 0
    Perc_Change = 0
    Ticker = ""
    
    
'Loop through tickers

    'Output Table
    Cells(1, 12).Value = "Ticker"
    Cells(1, 13).Value = "Quarterly Change"
    Cells(1, 14).Value = "Percent Change"
    Cells(1, 15).Value = "Total Stock Volume"
    
    'Continue loop while ticker cell is not empty
    While Not IsEmpty(Cells(Row, 1))

    
'Check if we are still within the same ticker
    'Check if ticker = Row behind
    If Cells(Row, 1).Value = Cells(Row - 1, 1).Value Then

        'add the volume to the total
        Volume = Volume + Cells(row,7).Value
    Else
        'this is last row of quarter, record end price
        End_Price = Cells(row,6).Value
        'and add the volume to the total
        Volume = Volume + Cells(row,7).Value
        'record ticker
        Ticker = Cells(row,1).Value
        
    
    
 



End Sub