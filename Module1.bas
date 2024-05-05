Attribute VB_Name = "Module1"
Sub SummaryStockDataAllQuarters()
    Dim quarterSheets As Variant
    quarterSheets = Array("Q1", "Q2", "Q3", "Q4")
    
    Dim sheetName As Variant
    For Each sheetName In quarterSheets
        If Not Evaluate("ISREF('" & sheetName & "'!A1)") Then
            MsgBox "Sheet " & sheetName & " does not exist!", vbExclamation
            Exit Sub
        End If
    Next sheetName
    
    Dim ws As Worksheet
    For Each sheetName In quarterSheets
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        ' Add headers for the summary columns
        ws.Cells(1, 9).Value = "Ticker Symbol"
        ws.Cells(1, 10).Value = "Total Quarterly Change"
        ws.Cells(1, 11).Value = "Total Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Assuming the first row has headers and data starts from the second row
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Create a dictionary to store unique tickers and their summaries
        Dim summary As Object
        Set summary = CreateObject("Scripting.Dictionary")
        
        ' Variables to hold the greatest values
        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        Dim greatestVolume As Double
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        
        ' Variables to hold the tickers for the greatest values
        Dim tickerIncrease As String
        Dim tickerDecrease As String
        Dim tickerVolume As String
        
        ' Loop through each stock
        Dim i As Long
        For i = 2 To lastRow
            Dim ticker As String
            Dim dateValue As Date
            Dim volume As Double
            ticker = ws.Cells(i, 1).Value
            dateValue = ws.Cells(i, 2).Value
            volume = CDbl(ws.Cells(i, 7).Value)
            
            ' Initialize the dictionary for the new ticker
            If Not summary.Exists(ticker) Then
                summary.Add ticker, CreateObject("Scripting.Dictionary")
                summary(ticker).Add "Opening Price", ws.Cells(i, 3).Value
                summary(ticker).Add "Volume", volume
                summary(ticker).Add "Start Date", dateValue
            Else
                ' Update the total volume
                summary(ticker)("Volume") = summary(ticker)("Volume") + volume
                ' Update the closing price if the current date is later than the stored date
                If dateValue > summary(ticker)("Start Date") Then
                    summary(ticker)("Closing Price") = ws.Cells(i, 6).Value
                    summary(ticker)("End Date") = dateValue
                End If
            End If
        Next i
        
        ' Calculate the quarterly change and percent change for each ticker
        Dim key As Variant
        For Each key In summary.Keys
            Dim openingPrice As Double
            Dim closingPrice As Double
            Dim totalVolume As Double
            Dim quarterlyChange As Double
            Dim percentChange As Double
            
            openingPrice = summary(key)("Opening Price")
            closingPrice = summary(key)("Closing Price")
            totalVolume = summary(key)("Volume")
            
            ' Calculate changes
            quarterlyChange = closingPrice - openingPrice
            If openingPrice <> 0 Then
                percentChange = (quarterlyChange / openingPrice) * 100
            Else
                percentChange = 0
            End If
            
            ' Update the dictionary with the calculated values
            summary(key)("Quarterly Change") = quarterlyChange
            summary(key)("Percent Change") = percentChange
            
            ' Check for greatest increase, decrease, and volume
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                tickerIncrease = key
            ElseIf percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                tickerDecrease = key
            End If
            
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                tickerVolume = key
            End If
        Next key
        
        ' Output the summary to the specified columns
        Dim j As Long
        j = 2 ' Start from the second row
        For Each key In summary.Keys
            ws.Cells(j, 9).Value = key
            ws.Cells(j, 10).Value = summary(key)("Quarterly Change")
            ws.Cells(j, 11).Value = Format(summary(key)("Percent Change"), "0.00") & "%"
            ws.Cells(j, 12).Value = summary(key)("Volume")
            j = j + 1
        Next key
        
        ' Output the greatest values
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = tickerIncrease
        ws.Cells(2, 17).Value = Format(greatestIncrease, "0.00") & "%"
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = tickerDecrease
        ws.Cells(3, 17).Value = Format(greatestDecrease, "0.00") & "%"
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = tickerVolume
        ws.Cells(4, 17).Value = greatestVolume
        
        ' Add headers for the greatest values
        ws.Cells(1, 15).Value = "Metric"
        ws.Cells(1, 16).Value = "Ticker Symbol"
        ws.Cells(1, 17).Value = "Value"
        
        ' AutoFit the columns for better readability
        ws.Columns("I:L").AutoFit
        ws.Columns("O:Q").AutoFit
    Next sheetName
End Sub

