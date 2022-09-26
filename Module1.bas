Attribute VB_Name = "Module1"


Sub Stock_Market()

    'Set initial variable for worksheets
    
    'Set initial variable for holding Ticker
    Dim Ticker As String
    
    'set initial variable for lastrow
    Dim lastrow As Long
    
    'Set initial variable for open price
    Dim open_price As Double
    
    'Set variable for close price
    Dim close_price As Double
    
    
    ' Set Initial variable for yearly change
    Dim Yearly_Change As Double
        Yearly_Change = 0
    'Set initial variable Percent Change
    Dim Percent_Change As Double
        Percent_Change = 0
    
    'Set Initial variable Total stock volume
    Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0
    
    'Keep track of location of each ticker symbol
    Dim Ticker_Symbol_Row As Integer
    Ticker_Symbol_Row = 2
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    open_price = ws.Cells(2, 3).Value
    
    close_price = ws.Cells(I, 6).Value
    
    
    
 ' Loop through all ticker Symbols
    
        For I = 2 To lastrow
    
        'check if we are still within the same Ticker symbol, if not...
            If ws.Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        
            'Set the Ticker Symbol
            Ticker = ws.Cells(I, 1).Value
            
            'Yearly change from open to close
            close_price = ws.Cells(I, 6).Value
            
            'Print the ticker symbol
            Range("K" & Ticker_Symbol_Row).Value = Ticker
            
            'Print Brand amount to summary table
            Yearly_Change = close_price - open_price
            
            Percent_Change = (Yearly_Change / open_price)
                If Percent_Change <= 0 Then
                    Percent_Change.Interior.ColorIndex = 3
                Else
                    Percent_Change.Interior.ColorIndex = 4
                End If
            
            Total_Stock_Volume = Application.WorksheetFunction.Sum(Range("G"))
            
            
            'Add one to the Ticker Symbol
            Ticker_Symbol_Row = Ticker_Symbol_Row + 1
            
            'reset the Yearly Change
            Yearly_Change = 0
             
        'If the cell uimmediately following a row is the same brand..
        Else
        
            'Add to the Brand Total
            Yearly_Change = Yearly_Change + Cells(I, 3).Value
            
    End If
        
    Next I
    
End Sub

    

    

 
�� 