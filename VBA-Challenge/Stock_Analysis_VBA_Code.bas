Attribute VB_Name = "sheet3_working"
Sub HWT()
Dim tick As String
Dim Volume As Double
Dim tablePoisiton As Integer
Dim finalRow
Dim firstTick As Single
Dim finalTick As Single


tableposition = 2
Volume = 0
tick = "A"
firstTick = Cells(2, 3)

    'Gets last row of sheet
    finalRow = Cells(Rows.Count, "A").End(xlUp).Row + 1
    'MsgBox (finalRow & ws.Name)
    
    'provides row headers
    Range("I1").Value = "Ticker symbol"
    Range("J1").Value = "Yearly change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock volume"
    
    
    For i = 2 To finalRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            finalTick = Cells(i, 6).Value
                    
                    'code below was for debugging purposes
                    'These comments below were to test the first and final tick values
                    'Range("M" & tableposition).Value = firstTick
                    'Range("N" & tableposition).Value = finalTick
            
            'Calculates yearly change Column
            yearlyChange = finalTick - firstTick
            Range("J" & tableposition).Value = yearlyChange
            
    
            If firstTick > 0 Then
                'Calculates yearly percent change column
                yearlyPercentChange = -1 * ((firstTick - finalTick) / firstTick)
                Range("K" & tableposition).Value = yearlyPercentChange
            End If
                    
        
            'gets ticket value
            tick = Cells(i, 1).Value
        
            'totals volume
            Volume = Volume + Cells(i, 7).Value
        
            'Paste ticker name in left hand column
            Range("I" & tableposition).Value = tick
        
            'Paste total volume in table
            Range("L" & tableposition).Value = Volume
        
            'add one to table position
            tableposition = tableposition + 1
        
            'reset volume value
             Volume = 0
        
          'if the next cell is also the same ticket, then add to total volume
        Else
            
            If tick <> Cells(i, 1) Then
                If Cells(i, 1) > 0 Then
                    tick = Cells(i, 1).Value
                    firstTick = Cells(i, 3).Value
                End If
            End If
            
            'add to volume total
            Volume = Volume + Cells(i, 7).Value
            
        End If
    Next i



End Sub

