Attribute VB_Name = "Module1"
Sub Stock_Challenge()

' Create variables
    Dim LastRow As Long
    Dim InfoTable As Long
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim StockVol As LongLong
    
' Create variables for Challenges
    Dim LastRow2 As Long
    Dim GreatIncTot As Double
    Dim GreatDecTot As Double
    Dim GreatVolTot As LongLong
    
    ' Challenge Part 2 - run script on every sheet
    For Each Sheet In Worksheets
        
        ' Set variables for summary table to place all Stocks and calculations into
        LastRow = Sheet.Cells(Rows.Count, 1).End(xlUp).Row
        InfoTable = 2
        
        '**Step 1 - need to sort the data by 2 different columns**
        'Source: https://trumpexcel.com/sort-data-vba/
        
        With Sheet.Sort
            .SortFields.Add Key:=Sheet.Range("A1"), Order:=xlAscending
            .SortFields.Add Key:=Sheet.Range("B1"), Order:=xlAscending
            .SetRange Sheet.Range("A1:G" & LastRow)
            .Header = xlYes
            .Apply
        End With
        
        'Create InfoTable headers
        Sheet.Range("I1").Value = "Ticker"
        Sheet.Range("J1").Value = "Yearly Change"
        Sheet.Range("K1").Value = "Percent Change"
        Sheet.Range("L1").Value = "Total Stock Volume"
        
        'Challenge Part 1 table headers and row names
        Sheet.Range("P1").Value = "Ticker"
        Sheet.Range("Q1").Value = "Value"
        Sheet.Range("O2").Value = "Greatest % Increase"
        Sheet.Range("O3").Value = "Greatest % Decrease"
        Sheet.Range("O4").Value = "Greatest Total Volume"
    
    ' Create InfoTable and fill in ticker value
    
        For i = 2 To LastRow
            
            'Find Opening price on earliest week for each stock
            If Sheet.Cells(i, 1).Value <> Sheet.Cells(i - 1, 1).Value Then
                OpenPrice = Sheet.Cells(i, 3).Value
            End If
        
            'Find Closing price on last week for each stock
            If Sheet.Cells(i, 1).Value <> Sheet.Cells(i + 1, 1).Value Then
                ' Show Stock name in table
                Sheet.Range("I" & InfoTable) = Sheet.Cells(i, 1).Value
                ClosePrice = Sheet.Cells(i, 6).Value
                ' Calculate Yearly Change
                YearlyChange = ClosePrice - OpenPrice
                Sheet.Range("J" & InfoTable).Value = YearlyChange
                
                ' Conditional Formatting for positive or negative yearly change
                
                    If YearlyChange > 0 Then
                        Sheet.Range("J" & InfoTable).Interior.ColorIndex = 4
                    ElseIf YearlyChange < 0 Then
                        Sheet.Range("J" & InfoTable).Interior.ColorIndex = 3
                    End If
                
                'Calculate Percent Change ( must specify OpenPrice above 0 to avoid division by 0)
                    If OpenPrice > 0 Then
                        PercentChange = (YearlyChange / OpenPrice) * 100
                        Sheet.Range("K" & InfoTable).Value = PercentChange & "%"
                    Else
                        PercentChange = 0
                        Sheet.Range("K" & InfoTable).Value = PercentChange & "%"
                    End If
                    
                'Calculate total stock volume
                StockVol = StockVol + Sheet.Cells(i, 7).Value
                Sheet.Range("L" & InfoTable).Value = StockVol
                'Reset StockVol counter for next stock
                StockVol = 0
                'Add new line to Summary table for next stock
                InfoTable = InfoTable + 1
                
            Else
                StockVol = StockVol + Sheet.Cells(i, 7).Value
                 
            End If
        
        
        Next i

        'Challenge Part 1
        LastRow2 = Sheet.Cells(Rows.Count, 11).End(xlUp).Row
       
       'Find highest percent increase value
        GreatIncTot = WorksheetFunction.Max(Sheet.Range("K2:K" & LastRow2))
        Sheet.Range("Q2").Value = GreatIncTot
        'Find highest percent decrease value
        Sheet.Range("Q2").NumberFormat = "0.00%"
        GreatDecTot = WorksheetFunction.Min(Sheet.Range("K2:K" & LastRow2))
        Sheet.Range("Q3").Value = GreatDecTot
        Sheet.Range("Q3").NumberFormat = "0.00%"
        'Find largest stock volume value
        GreatVolTot = WorksheetFunction.Max(Sheet.Range("L2:L" & LastRow2))
        Sheet.Range("Q4").Value = GreatVolTot
            
        'Search summary table for the above values and print corresponding stock name
        For i = 2 To LastRow2
            If Sheet.Cells(i, 11).Value = GreatIncTot Then
                Sheet.Range("P2") = Sheet.Cells(i, 9).Value
            ElseIf Sheet.Cells(i, 11).Value = GreatDecTot Then
                Sheet.Range("P3") = Sheet.Cells(i, 9).Value
            ElseIf Sheet.Cells(i, 12).Value = GreatVolTot Then
                Sheet.Range("P4") = Sheet.Cells(i, 9).Value
            End If
            
         Next i
    
    Next Sheet
        
End Sub

