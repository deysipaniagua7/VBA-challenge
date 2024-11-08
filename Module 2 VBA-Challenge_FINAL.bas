Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data()
    Dim ws As Worksheet
    Dim i, LastRow, Summary_Table_Row As Long
    Dim Ticker As String
    Dim Quarterly_Change, Total_Percentage_Change, Total_Stock_Volume As Double
  
    'List Worksheets to loop
    Dim sheetNames() As Variant
    sheetNames = Array("Q1", "Q2", "Q3", "Q4")
        
    'Loop through worksheets
    For Each ws In ThisWorkbook.Worksheets
        If Not IsError(Application.Match(ws.Name, sheetNames, 0)) Then
            ws.Activate
        
            'Set variable for holding quarterly change and total percentage from open to close price from start to end of quarter AND total stock volume
            Quarterly_Change = 0
            Total_Percentage_Change = 0
            Total_Stock_Volume = 0
            
            'Set variables for holding and tracking greatest values
            Dim Greatest_Increase, Greatest_Decrease, Greatest_Volume As Double
            Dim Greatest_Increase_Ticker, Greatest_Decrease_Ticker, Greatest_Volume_Ticker As String
            Greatest_Increase = 0
            Greatest_Decrease = 0
            Greatest_Volume = 0
            
            'Keep track of the location for each Quarterly_Change, Total_Percentage_Change, Total_Stock Volume
            Summary_Table_Row = 2
        
            'Set headers after setting variables above
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Quarterly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
                
            'Get the row number of the last row with data
            LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
                'Set loop
                For i = 2 To LastRow
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                        Ticker = ws.Cells(i, 1).Value
                        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                        Quarterly_Change = ws.Cells(i, 6) - ws.Cells(i, 3).Value
                
                        'Color code variable values for Quarterly_Change (column j)
                        If Quarterly_Change > 0 Then
                            ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                        ElseIf Quarterly_Change < 0 Then
                            ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                        Else
                            ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 2
                        End If
                        'Calculate Percentage_Change
                        If ws.Cells(i, 3).Value <> 0 Then
                            Total_Percentage_Change = Quarterly_Change / ws.Cells(i, 3).Value
                        Else
                            Total_Percentage_Change = 0
                        End If
                
                        'Display values for each column in summary table
                        ws.Cells(Summary_Table_Row, 9).Value = Ticker
                        ws.Cells(Summary_Table_Row, 10).Value = Quarterly_Change
                        ws.Cells(Summary_Table_Row, 10).NumberFormat = "0.00"
                        ws.Cells(Summary_Table_Row, 11).Value = Total_Percentage_Change
                        ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
                        ws.Cells(Summary_Table_Row, 12).Value = Total_Stock_Volume
                    
                        'Track greatest values
                        If Total_Percentage_Change > Greatest_Increase Then
                            Greatest_Increase = Total_Percentage_Change
                            Greatest_Increase_Ticker = Ticker
                        End If
                        If Total_Percentage_Change < Greatest_Increase Then
                            Greatest_Decrease = Total_Percentage_Change
                            Greatest_Decrease_Ticker = Ticker
                        End If
                        If Total_Stock_Volume > Greatest_Voume Then
                            Greatest_Volume = Total_Stock_Volume
                            Greatest_Volume_Ticker = Ticker
                        End If
                
                        'Go to the next row in summary table
                        Summary_Table_Row = Summary_Table_Row + 1
                        Total_Stock_Volume = 0
                    End If
                Next i
            End If
    Next ws
    
    'Display headers and greatest values summary
        With ThisWorkbook.Worksheets(1)
        .Range("P1").Value = "Ticker"
        .Range("Q1").Value = "Value"
        .Range("O2").Value = "Greatest % Increase"
        .Range("O3").Value = "Greatest % Decrease"
        .Range("O4").Value = "Greatest Total Volume"
        .Range("P2").Value = Greatest_Increase_Ticker
        .Range("P3").Value = Greatest_Decrease_Ticker
        .Range("P4").Value = Greatest_Volume_Ticker
        .Range("Q2").Value = Greatest_Increase
        .Range("Q2").NumberFormat = "0.00%"
        .Range("Q3").Value = Greatest_Decrease
        .Range("Q3").NumberFormat = "0.00%"
        .Range("Q4").Value = Greatest_Volume
    End With
End Sub
           
         
