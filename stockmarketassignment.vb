Sub stockmarketassignment():

    'Calls for all worksheets to run in the for loop
    For Each ws In Worksheets

        'Defines each needed header 
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        'Define each needed variable 
        Dim TN As String
    'TN is saved as a string since it has no numeric value and is just a series of characters
        Dim LR As Long
        Dim TTV As Double
        TTV = 0
        Dim STR As Long
        STR = 2
        Dim Y_Open As Double
        Dim Y_Close As Double
        Dim Y_Change As Double
        Dim P_Amount As Long
        P_Amount = 2
        Dim PercentChange As Double
        Dim LastRowValue As Long
    'The rest of the variables are saved as Long and Double variables since some contain integers that are too large, and some numbers need to include decimal places

        'Finds out the location of the last row
     LR = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LR

            'Adds value to ticker total volume
            TTV = TTV + ws.Cells(i, 7).Value
            'Makes sure we are under the same ticker name if not moves onto the next value
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'Creates ticker name
                TN = ws.Cells(i, 1).Value
                'Puts ticker name in the summary table
                ws.Range("I" & STR).Value = TN
                'Puts ticker total amount in the summary table
                ws.Range("L" & STR).Value = TTV
                'Resets Ticker Total to zero
                TTV = 0

                'Sets up Yearly Open, Yearly Close and Yearly Change Name
                Y_Open = ws.Range("C" & P_Amount)
                Y_Close = ws.Range("F" & i)
                Y_Change = Y_Close - Y_Open
                ws.Range("J" & STR).Value = Y_Change

                'This code determines the percentage of change
                If Y_Open = 0 Then
                    PercentChange = 0
                Else
                    Y_Open = ws.Range("C" & P_Amount)
                    PercentChange = Y_Change / Y_Open
                End If
                'Formats the numbers to include two decimal places and to include the % symbol
                ws.Range("K" & STR).NumberFormat = "0.00%"
                ws.Range("K" & STR).Value = PercentChange

                'Conditional statement used to highlight the positive(green) and negative(red) values
                If ws.Range("J" & STR).Value >= 0 Then
                    ws.Range("J" & STR).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & STR).Interior.ColorIndex = 3
                End If
            
                'Add One To The Summary Table Row
                STR = STR + 1
                P_Amount = i + 1
                End If
        Next i
    Next ws
End Sub
