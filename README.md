# VBA
VBA Home Work
# Code and Excel will be uploaded here
the excel is stored here - https://drive.google.com/open?id=12qnkwzoq1w0Ta86oj9wgcmrDtDNEznQP
here is the code:
Sub Ticker_Total()

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------

    Dim totcounter As Long
    Dim tickertotal As Double
    Dim LastRow As Long
    Dim openprice As Double
    Dim closeprice As Double
    Dim tickerswitch As Integer
    Dim yearlychange As Double
    Dim percentchange As Double
    
' Variables for % Greatest % Lowest and highest tikcer Volume
    Dim max As Double
    Dim min As Double
    Dim minrow As Integer
    Dim maxrow As Integer
    Dim maxvolrow As Integer
    Dim maxvol As Double

    For Each ws In Worksheets
        totcounter = 2
        tickertotal = 0
        tickerswitch = 1
        
        ws.Cells(totcounter - 1, 9).Value = "Ticker"
        ws.Cells(totcounter - 1, 10).Value = "Yearly Change"
        ws.Cells(totcounter - 1, 11).Value = "Percent Change"
        ws.Cells(totcounter - 1, 12).Value = "Total Stock Volume"

        ws.Cells(totcounter - 1, 16).Value = "Ticker"
        ws.Cells(totcounter - 1, 17).Value = "Value"

        ' Determine the Last Row
        ' Get the active last row in ticker column
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow
        ' when the ticker changes sum up the volume and reset the counters
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                tickertotal = tickertotal + ws.Cells(i, 7).Value
                ws.Cells(totcounter, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(totcounter, 12).Value = tickertotal
                tickertotal = 0
                
                closeprice = ws.Cells(i, 6).Value
                yearlychange = closeprice - openprice
                ws.Cells(totcounter, 10).Value = yearlychange
                percentchange = (closeprice - openprice) * 100 / (openprice)
                ws.Cells(totcounter, 11).Value = percentchange
                'Conditional formatting
                If yearlychange > 0 Then
                    With ws.Cells(totcounter, 10)
                    'With Cells(totcounter, 10).FormatConditions.Add(xlCellValue, xlLess, "=0")
                 .Interior.Color = vbGreen
                 .Font.Color = vbBlack
                 End With
                Else
                    ' (2) Highlight defined good margin as green values
                    With ws.Cells(totcounter, 10)
                        .Interior.Color = vbRed
                        .Font.Color = vbBlack
                    End With
                End If
                tickerswitch = 1
                totcounter = totcounter + 1
                
            Else
                tickertotal = tickertotal + ws.Cells(i, 7).Value
                If tickerswitch = 1 And ws.Cells(i, 3) <> 0 Then
                   openprice = ws.Cells(i, 3).Value
                   tickerswitch = 0
                End If
            End If
        
        Next i
        
        LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        max = 0 'set "inital" max
        min = 0 'set "inital" min
        maxvol = 0 'set "inital" max
        
            For i = 2 To LastRow
                'loop through values
                If ws.Cells(i, 12).Value <> 0 Then
                        ' locate the highest volume and save the corresonding row values
                        If ws.Cells(i, 12).Value > maxvol And ws.Cells(i, 12).Value <> 0 Then 'if a value is larger than the old max,
                             maxvol = ws.Cells(i, 12).Value ' store it as the new max!
                             maxvolrow = i
                        End If
        
                        ' locate the highest % increase and save the corresonding row values
                        If ws.Cells(i, 11).Value > max And ws.Cells(i, 12).Value <> 0 Then 'if a value is larger than the old max,
                             max = ws.Cells(i, 11).Value ' store it as the new max!
                             maxrow = i
                        End If
        
                        ' locate the highest % decrease and save the corresonding row values
                        If ws.Cells(i, 11).Value < min And ws.Cells(i, 12).Value <> 0 Then 'if a value is less than the old min ,
                             min = ws.Cells(i, 11).Value ' store it as the new min!
                             minrow = i
                        End If
            End If
            Next i

            ' Move the % high increase values
        ws.Cells(2, 17).Value = max
        ws.Cells(2, 16).Value = ws.Cells(maxrow, 9).Value
        ws.Cells(2, 15).Value = "Greatest % Increase"

        ' Move the % high decrease values
        ws.Cells(3, 17).Value = min
        ws.Cells(3, 16).Value = ws.Cells(minrow, 9).Value
        ws.Cells(3, 15).Value = "Greatest % Decrease"

        ' Move the Max Volume values
        ws.Cells(4, 17).Value = maxvol
        ws.Cells(4, 16).Value = ws.Cells(maxvolrow, 9).Value
        ws.Cells(4, 15).Value = "Greatest % Volume"
 
        ' MsgBox (ws.Name)
        'Exit For
        
     Next ws
   
End Sub



