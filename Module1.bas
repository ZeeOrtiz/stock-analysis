Attribute VB_Name = "Module1"
Sub MacroCheck()

    Dim testmessage As String
    
    testmessage = "Hello World"

    MsgBox (testmessage)

End Sub

Sub DQAnalysis()
           
    Worksheets("DQ Analysis").Activate

    
    Range("A1").Value = "DAQO (Ticker: DQ)"

    
    
    'Create a header row
    
    Cells(3, 1).Value = "Year"
    
    Cells(3, 2).Value = "Total Daily Volume"
    
    Cells(3, 3).Value = "Return"
     
    
    Worksheets("2018").Activate
    
    
    'set initial volume to zero
    totalVolume = 0
    
    
    Dim startingPrice As Double
    
    Dim endingPrice As Double
    
    
    'Establish the number of rows to loop over
    
    rowStart = 2
    
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    
   
   'loop over all rows
    
    For i = rowStart To rowEnd
    
    
    
        If Cells(i, 1).Value = "DQ" Then
    
    
                'increase totalVolume if ticker is "DQ"
    
                totalVolume = totalVolume + Cells(i, 8).Value
        End If
    
    
    If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
    
        startingPrice = Cells(i, 6).Value
        
        
     End If

    If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
    
        endingPrice = Cells(i, 6).Value
    
        End If
    
    
    Next i
    
      
    
    Worksheets("DQ Analysis").Activate
    
    Cells(4, 1).Value = 2018
    
    Cells(4, 2).Value = totalVolume
    
    Cells(4, 3).Value = (endingPrice / startingPrice) - 1
    

End Sub

Sub AllStocksAnalysis()
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"

    
    
    'Create a header row
    
    Cells(3, 1).Value = "Ticker"
    
    Cells(3, 2).Value = "Total Daily Volume"
    
    Cells(3, 3).Value = "Return"
     
 'initialize array of all tickers
 
      
    Dim tickers(11) As String
        
        tickers(0) = "AY"
        tickers(1) = "CSIQ"
        tickers(2) = "DQ"
        tickers(3) = "ENPH"
        tickers(4) = "FSLR"
        tickers(5) = "HASI"
        tickers(6) = "JKS"
        tickers(7) = "RUN"
        tickers(8) = "SEDG"
        tickers(9) = "SPWR"
        tickers(10) = "TERP"
        tickers(11) = "VSLR"
        
    'start and end price
    Dim startingPrice As Single
    Dim endingPrice As Single
    
Worksheets("2018").Activate
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
 
           Worksheets("2018").Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
 
 If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If
If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
Next j

Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i
  
    
    
End Sub

Sub formatAllStocksAnalysisTable()
'Formatting
Worksheets("All Stocks Analysis").Activate

Range("A3:C3").Font.Bold = True

Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous

Range("A3:C3").Font.ThemeFont = xlThemeFontMajor

Range("B4:B15").NumberFormat = "#,##0"

Range("C4:C15").NumberFormat = "0.0%"

Range("B4:B15").NumberFormat = "$#,##0.00"

Columns("B").AutoFit

If Cells(4, 3) > 0 Then
    'Color the cell green
    Cells(4, 3).Interior.Color = vbGreen
ElseIf Cells(4, 3) < 0 Then

    'Color the cell red
    Cells(4, 3).Interior.Color = vbRed

Else
    'Clear the cell color
    Cells(4, 3).Interior.Color = xlNone

End If

    dataRowStart = 4
    datarowEnd = 15
    For i = dataRowStart To datarowEnd

        If Cells(i, 3) > 0 Then

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        ElseIf Cells(i, 3) < 0 Then

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If

    Next i

End Sub

Sub ClearWorksheet()
  
    Cells.ClearContents
 
End Sub


Sub AllStocksAnalysis1()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

       startTime = Timer
Next i

    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

 End Sub
