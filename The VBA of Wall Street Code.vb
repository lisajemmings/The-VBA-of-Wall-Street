Sub VBAHomework()
    For Each ws In Worksheets
        ws.Activate
        Call CalculateSummary
    Next ws
End Sub

Sub CalculateSummary()

' Setting Variables
Dim Ticker As String
Dim Volume As Double
Dim Open_Price As Double
Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim LastRow As Long
Dim Ticker_Row As Long
Dim Row As Long
Dim Column As Integer

' Setting Values
Volume = 0
Ticker_Row = 2
Column = 1

' Setting Headers
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

Range("I:L").EntireColumn.AutoFit

LastRow = Cells(Rows.Count, 1).End(xlUp).Row

Open_Price = Cells(2, Column + 2).Value

    For Row = 2 To LastRow

        Volume = Volume + Cells(Row, Column + 6).Value
     
        If Cells(Row + 1, Column).Value <> Cells(Row, Column).Value Then
        
            ' Populating Ticker Column
            TickerSymbol = Cells(Row, Column).Value
            Range("I" & Ticker_Row).Value = TickerSymbol
    
            'Populating Yearly Change Column
            Close_Price = Cells(Row, Column + 5).Value
            Yearly_Change = Close_Price - Open_Price
            Range("J" & Ticker_Row).Value = Yearly_Change
                            
            'Populating Percent Change Column - Avoiding Division by 0
            
                If Open_Price = 0 Then
                
                    Percent_Change = 0
                    
                Else
                    
                    Percent_Change = Yearly_Change / Open_Price
                    Range("K" & Ticker_Row).Value = Percent_Change
                    Range("K" & Ticker_Row).NumberFormat = "0.00%"
                                       
                End If
            
            'Populating Total Stock Volume Column
            Range("L" & Ticker_Row).Value = Volume

            Ticker_Row = Ticker_Row + 1
            
            ' Reset
            Volume = 0
            Open_Price = Cells(Row + 1, Column + 2)
    
        End If
    
    Next Row
    
LastRow = Cells(Rows.Count, 9).End(xlUp).Row

    For Row = 2 To LastRow
        
        If Cells(Row, Column + 9).Value > 0 Then
        
            Cells(Row, Column + 9).Interior.ColorIndex = 10
            
        Else
        
            Cells(Row, Column + 9).Interior.ColorIndex = 3
            
    End If

    Next Row
    
Debug.Print ActiveSheet.Name

End Sub

