Attribute VB_Name = "Module1"
Sub ticker()
    
        Dim ws As Worksheet
        Dim t_count As Long
        Dim volume As Double
        Dim last_row As Long
        Dim yr_change As Double
        Dim p_change As Double
        Dim output_range As Integer
        Dim myrange As Range
        Dim lrow As Long
        
    
    
            output_range = 2
    
    
    For Each ws In Worksheets
    
            'set headers
            ws.Cells(1, 12).Value = "Ticker"
            ws.Cells(1, 13).Value = "Yearly Change"
            ws.Cells(1, 14).Value = "Percentage Change"
            ws.Cells(1, 15).Value = "Total Volume"
            ws.Cells(2, 18).Value = "Greatest % Increase"
            ws.Cells(3, 18).Value = "Greatest % Decrease"
            ws.Cells(4, 18).Value = "Greatest Total Volume"
            Columns(18).AutoFit
            ws.Cells(1, 19).Value = "Ticker"
            ws.Cells(1, 20).Value = "Value"
            
            last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            For i = 2 To last_row
            
                If ws.Cells(i + 1, 1) = ws.Cells(i, 1) Then
                    t_count = t_count + 1
                    volume = volume + ws.Cells(i, 7)
                    
                    If t_count = 1 Then
                        open_value = ws.Cells(i, 3)
                    Else
                End If
            
            Else
                ws.Cells(output_range, 12) = ws.Cells(i, 1)
                ws.Cells(output_range, 15) = volume
            
                close_value = ws.Cells(i, 6)
            
            If open_value <> 0 Then
                yr_change = close_value - open_value
                p_change = (close_value - open_value) / open_value
                
            Else
                yr_change = 0
                p_change = 0
                
            End If
            
            ws.Cells(output_range, 13).Value = yr_change
            ws.Cells(output_range, 14).Value = p_change
           
        'conditional formatting
            If ws.Cells(output_range, 13) < 0 Then
                    ws.Cells(output_range, 13).Interior.ColorIndex = 3
                    ws.Cells(output_range, 14).NumberFormat = "0.00%"
            Else
                    ws.Cells(output_range, 13).Interior.ColorIndex = 10
                    ws.Cells(output_range, 14).NumberFormat = "0.00%"
            End If
            
        'reset
            t_count = 0
            volume = 0
            output_range = output_range + 1
         
            End If
        Next i
         
        'output_range reset
        output_range = 2
        'bonus
                
        maxvolume = Application.WorksheetFunction.Max(Range("O:O"))
        maxvalue = Application.WorksheetFunction.Max(Range("N:N"))
        minvalue = Application.WorksheetFunction.Min(Range("N:N"))
        ws.Cells(2, 20).Value = maxvalue
        ws.Cells(3, 20).Value = minvalue
        ws.Cells(4, 20).Value = maxvolume
        ws.Cells(2, 20).NumberFormat = "0.00%"
        ws.Cells(3, 20).NumberFormat = "0.00%"
        
        
        For j = 2 To last_row
            If ws.Cells(j, 14) = maxvalue Then
                    ws.Cells(2, 19).Value = ws.Cells(j, 12).Value
                    
                End If
            If ws.Cells(j, 14) = minvalue Then
                    ws.Cells(3, 19).Value = ws.Cells(j, 12).Value
                    
                End If
            If ws.Cells(j, 15) = maxvolume Then
                ws.Cells(4, 19).Value = ws.Cells(j, 12).Value
                    
                End If
        Next j
        Columns(20).AutoFit
      
    Next ws
End Sub
