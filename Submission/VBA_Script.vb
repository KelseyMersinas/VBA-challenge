Sub open_to_close():
'declared values
Dim i As Long
Dim k As Integer
Dim j As Integer
Dim last_row As Long
Dim ticker As String
Dim start As Double
Dim closing As Double
Dim change As Double
Dim per_change As Double
Dim volume_cell As Double
Dim vol_total As Double
Dim CurSheet As Worksheet
Dim highestpct As Double
Dim lowestpct As Double
Dim highestVol As Double
Dim tik_high As String
Dim tik_low As String
Dim tik_vol As String

For Each CurSheet In ThisWorkbook.Worksheets 'loop though each worksheet
    'reseting values at the top of each sheet
    k = 2
    last_row = CurSheet.Cells(CurSheet.Rows.Count, 1).End(xlUp).Row
    highestpct = 0
    lowestpct = 0
    highestVol = 0
    'add in column headers for each sheet
    CurSheet.Cells(1, 9).Value = "Ticker"
    CurSheet.Cells(1, 10).Value = "Quarterly Change"
    CurSheet.Cells(1, 11).Value = "Percent Change"
    CurSheet.Cells(1, 12).Value = "Total Stock Value"
    CurSheet.Cells(1, 16).Value = "Ticker"
    CurSheet.Cells(1, 17).Value = "Value"
    CurSheet.Cells(2, 15).Value = "Greatest % Increase"
    CurSheet.Cells(3, 15).Value = "Greatest % Decrease"
    CurSheet.Cells(4, 15).Value = "Greatest Total Volume"
    
    'loop though column A
    For i = 2 To last_row
        ticker = CurSheet.Cells(i, 1).Value
        volume_cell = CurSheet.Cells(i, 7).Value
        
        'if fisrt time seeing new ticker
        If (CurSheet.Cells(i - 1, 1).Value <> ticker) Then
            start = CurSheet.Cells(i, 3).Value
            vol_total = volume_cell
            
        'if last line of a ticker
        ElseIf (CurSheet.Cells(i + 1, 1).Value <> ticker) Then
            'get value for change between them
            change = (CurSheet.Cells(i, 6).Value - start)
            'use change value to get percentage
            'technically stocks cannot be 0 dollars but- check for division by 0
            If start <> 0 Then
                'don't have to *100 since displaying them as a % will do that for us
                per_change = (change / start)
            Else
                per_change = 0
            End If
            
            vol_total = vol_total + volume_cell
         
            'display results of i loop in the columns I:L
            CurSheet.Cells(k, 9).Value = ticker
            CurSheet.Cells(k, 10).Value = change
            CurSheet.Cells(k, 11).Value = per_change
            CurSheet.Cells(k, 12).Value = vol_total
            
            'conditional formating for quarterly chnage
            If (CurSheet.Cells(k, 10).Value > 0) Then
                CurSheet.Cells(k, 10).Interior.ColorIndex = 4
                        
            ElseIf (CurSheet.Cells(k, 10).Value < 0) Then
                CurSheet.Cells(k, 10).Interior.ColorIndex = 3
            End If
                
            k = k + 1
        'otherwise keep the volume total running
        Else
            vol_total = vol_total + volume_cell
        
        End If
        
     
    Next i
    'loop through those results for highest percent change
    For j = 2 To (CurSheet.Cells(CurSheet.Rows.Count, 12).End(xlUp).Row)
        'if the next cell has a higher number
        If (CurSheet.Cells(j, 11).Value > highestpct) Then
            highestpct = CurSheet.Cells(j, 11).Value
            tik_high = CurSheet.Cells(j, 9).Value
        'if the cell has a lower number
        ElseIf (CurSheet.Cells(j, 11).Value < lowestpct) Then
            lowestpct = CurSheet.Cells(j, 11).Value
            tik_low = CurSheet.Cells(j, 9).Value
                           
        End If
        'finding the highest volumn
        If CurSheet.Cells(j, 12).Value > highestVol Then
            highestVol = CurSheet.Cells(j, 12).Value
            tik_vol = CurSheet.Cells(j, 9).Value
                        
        End If
                    
        Next j
        
    'display the results of the higest and lowest values in columns 0:Q
    CurSheet.Cells(2, 17).Value = highestpct
    CurSheet.Cells(3, 17).Value = lowestpct
    CurSheet.Cells(4, 17).Value = highestVol
    CurSheet.Cells(2, 16).Value = tik_high
    CurSheet.Cells(3, 16).Value = tik_low
    CurSheet.Cells(4, 16).Value = tik_vol
    
    
Next CurSheet
End Sub


