Sub Multiple_year_stock_data():

'Creating a Variable for worksheets in the currect workbook and to store the info of row_count/last_row(LR) in each sheet
Dim WS As Worksheet
Dim LR As Long

'Looping in worksheets
For Each WS In ActiveWorkbook.Worksheets

    LR = WS.Range("A1").CurrentRegion.Rows.Count 'to count the last row in each worksheet
    
    'Creating columns and their headings for the results obtained from the raw data
    
    WS.Cells(1, 9).Value = "Ticker"
    WS.Cells(1, 10).Value = "Yearly Change"
    WS.Cells(1, 11).Value = "Percent Change"
    WS.Cells(1, 12).Value = "Total Stock Volume"
        
    'Declaring variables to hold the values; YC=yearly change, total=total stock volume
    Dim I, j As Integer
    Dim total As LongLong
    Dim YC As Double
    Dim open_value As Double
    Dim Close_value As Double
    Dim percent_change As Double
    
    j = 1
    
    open_value = WS.Cells(2, 3).Value
    
    'Second loop to go through all the values in raw data and to summarize the results in the new columns for each stock
    For I = 2 To LR
    
        'summarizing the data for each stock
        If WS.Cells(I, 1).Value <> WS.Cells(I + 1, 1).Value Then
            j = j + 1
                
                WS.Cells(j, 9) = WS.Cells(I, 1).Value
                
                total = total + WS.Cells(I, 7).Value
                WS.Cells(j, 12) = total
                
                Close_value = WS.Cells(I, 6)
                YC = Close_value - open_value
                WS.Cells(j, 10) = YC
                 
                    'conditional formatting of the cells based on yearly change
                    If YC > 0 Then
                        WS.Cells(j, 10).Interior.ColorIndex = 4
                        
                        ElseIf YC < 0 Then
                        WS.Cells(j, 10).Interior.ColorIndex = 3
                     
                    End If
                
                percent_change = (YC / open_value) * 100
                WS.Cells(j, 11).Value = (Str(percent_change) + "%")
                
                    'conditional formatting of the cells based on % change
                    If percent_change > 0 Then
                        WS.Cells(j, 11).Font.ColorIndex = 10
                        
                        ElseIf percent_change < 0 Then
                        WS.Cells(j, 11).Font.ColorIndex = 3
                    
                    End If
                    
                open_value = WS.Cells(I + 1, 3)
                total = 0 'reassigning the value to the total, to start from scratch to count for new stock volume
            
        
        ElseIf WS.Cells(I, 1).Value = WS.Cells(I + 1, 1).Value Then
            total = total + WS.Cells(I, 7).Value 'Keep on adding the number of stocks to the total (stock volume) until the stock name remains same
            
        End If
        
    Next I
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
    'Creating columns and their headings summarize the maximum and minimum %change and to find the greatest total volume(GTV)of the stock
    WS.Cells(1, 15).Value = "Ticker"
    WS.Cells(1, 16).Value = "Value"
    WS.Cells(2, 14).Value = "Greatest % increase"
    WS.Cells(3, 14).Value = "Greatest % decrease"
    WS.Cells(4, 14).Value = "Greatest Total Volume"

        'Declaring varibles to finding the maximum(Max_PCR) and minimum (Min_PCR) %change and to find the greatest total volume(GTV)of the stock
        Dim Max_PC, Min_PC As Double
        Dim LR_Summary As Integer 'LR_Summary is to find the last row of the summary table
        Dim r As Integer
        Dim GTV As LongLong
        
        LR_Summary = WS.Cells(Rows.Count, 11).End(xlUp).Row
        MsgBox (LR_Summary)
        
        'Setting the value 0, at the start of analysis
        Max_PC = 0
        Min_PC = 0
        GTV = 0
        
    'Looping through summary table to find the maximum and minimum values
    For r = 2 To LR_Summary
    
        'First IF, to find the maximum percent change (Max_PC)
        If WS.Cells(r, 11).Value > Max_PC Then
                
                Max_PC = WS.Cells(r, 11).Value
                WS.Cells(2, 15).Value = WS.Cells(r, 9).Value
                WS.Cells(2, 16).Value = (Str(Max_PC * 100) + "%")
                
                Else
                Max_PC = Max_PC
                
        
        End If
        
        'Second IF, to find the minimum percent change (Max_PC)
        If WS.Cells(r, 11).Value < Min_PC Then
                
                Min_PC = WS.Cells(r, 11).Value
                WS.Cells(3, 15).Value = WS.Cells(r, 9).Value
                WS.Cells(3, 16).Value = (Str(Min_PC * 100) + "%")
                
                Else
                Min_PC = Min_PC
                
        
        End If
        
        'Third IF, to find the greatest total stock Volume (GTV)
        If WS.Cells(r, 12).Value > GTV Then
                
                GTV = WS.Cells(r, 12).Value
                WS.Cells(4, 15).Value = WS.Cells(r, 9).Value
                WS.Cells(4, 16).Value = GTV
                
                Else
                GTV = GTV
        
        End If
            
        
    Next r
    
Next WS
        
End Sub