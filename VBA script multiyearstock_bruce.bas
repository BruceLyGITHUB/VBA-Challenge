Attribute VB_Name = "Module1"
Sub multiyearstock():

    For Each ws In Worksheets
    
        'using dim to declare variable types
        Dim WorksheetName As String
        
        Dim i As Long
        
        Dim q As Long
        
        Dim TickC As Long
        
        Dim LastRowI As Long
        
        Dim LastRowQ As Long
        
        Dim Percent As Double
        
        Dim GreatInc As Double
        
        Dim GreatDec As Double
        
        Dim GreatVol As Double
        
            WorksheetName = ws.Name
        
                'column headers
                ws.Cells(1, 9).Value = "Ticker"
                ws.Cells(1, 10).Value = "Yearly change"
                ws.Cells(1, 11).Value = "Percent change"
                ws.Cells(1, 12).Value = "Total stock volume"
                ws.Cells(1, 16).Value = "Ticker"
                ws.Cells(1, 17).Value = "Value"
                ws.Cells(2, 15).Value = "Greatest % increase"
                ws.Cells(3, 15).Value = "Greatest % decrease"
                ws.Cells(4, 15).Value = "Greatest total volume"
                
                
        'yearly change column
        TickC = 2
        
        q = 2
        
        LastRowQ = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            For i = 2 To LastRowQ
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ws.Cells(TickC, 9).Value = ws.Cells(i, 1).Value
                
                ws.Cells(TickC, 10).Value = ws.Cells(i, 6).Value - ws.Cells(q, 3).Value
                
                    'using conditional format
                    If ws.Cells(TickC, 10).Value > 0 Then
                    
                    'green color
                    ws.Cells(TickC, 10).Interior.ColorIndex = 4
                
                    Else
                    
                     'red color
                    ws.Cells(TickC, 10).Interior.ColorIndex = 3
                
                    End If
                    
                    'percent change column
                    If ws.Cells(q, 3).Value <> 0 Then
                    Percent = ((ws.Cells(i, 6).Value - ws.Cells(q, 3).Value) / ws.Cells(q, 3).Value)
                    
                    ws.Cells(TickC, 11).Value = Format(Percent, "Percent")
                    
                    Else
                    
                    ws.Cells(TickC, 11).Value = Format(0, "Percent")
                    
                    End If
                    
        'Total volume column
        ws.Cells(TickC, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(q, 7), ws.Cells(i, 7)))
                
            TickC = TickC + 1
                
            q = i + 1
                
        End If
            
    Next i
            
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'hard solution(bonus)
        GreatVol = ws.Cells(2, 12).Value
        GreatInc = ws.Cells(2, 11).Value
        GreatDec = ws.Cells(2, 11).Value
        
            For i = 2 To LastRowI
            
                If ws.Cells(i, 12).Value > GreatVol Then
                GreatVol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatVol = GreatVol
                
                End If
                
                If ws.Cells(i, 11).Value > GreatInc Then
                GreatInc = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatInc = GreatInc
                
                End If
                
                If ws.Cells(i, 11).Value < GreatDec Then
                GreatDec = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatDec = GreatDec
                
                End If
                
            ws.Cells(2, 17).Value = Format(GreatInc, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDec, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
            
            Next i
            
    Next ws
        
End Sub
