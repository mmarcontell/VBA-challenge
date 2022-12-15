Attribute VB_Name = "Module1"
Sub stocksChange():

    Dim i, r As Double
    Dim ticket_name As String
    
    Dim ticket_volume As Double
    ticket_volume = 0
    
    Dim summary_table_row As Double
    summary_table_row = 2
    
    Dim open_row As Double
    open_row = 2
    
    Dim open_value As Double
    Dim close_value As Double
    
    Dim max_percent As Double
    Dim min_percent As Double
    Dim max_vol As Double
    Dim big_ticket As String
    Dim low_ticket As String
    Dim vol_ticket As String
    Range("K2") = max_percent
    Range("K2") = min_percent
    Range("L2") = max_vol
    
    
    Dim last_row As Double
    last_row = Cells(Rows.count, 1).End(xlUp).Row

    
'name the columns
  Range("I1,P1") = "Ticker"
  Range("J1") = "Yearly Change"
  Range("K1") = "Percent Change"
  Range("L1") = "Total Stock Volume"
  Range("Q1") = "Value"
  Range("O2") = "Greatest % Increase"
  Range("O3") = "Greatest % Decrease"
  Range("O4") = "Greatest Total Volume"
  
    
' loop through all the stocks for one year
    For i = 2 To last_row
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticket_name = Cells(i, 1).Value
            ticket_volume = ticket_volume + Cells(i, 7).Value
            Range("I" & summary_table_row).Value = ticket_name
            Range("L" & summary_table_row).Value = ticket_volume
            open_value = Cells(open_row, 3)
            close_value = Cells(i, 6)
            Range("J" & summary_table_row).Value = close_value - open_value
            Range("K" & summary_table_row).Value = ((close_value - open_value) / open_value)
            open_row = i + 1
            summary_table_row = summary_table_row + 1
            ticket_volume = 0
            
        Else
            ticket_volume = ticket_volume + Cells(i, 7).Value
            
        End If
        
 
    Next i
    
Dim shorter_last_row As Integer
shorter_last_row = Cells(Rows.count, 11).End(xlUp).Row
    
    For r = 2 To shorter_last_row
    
'returning max
        If Cells(r + 1, 11).Value > max_percent Then
            max_percent = Cells(r + 1, 11)
            big_ticket = Cells(r + 1, 9)
        Else
            Range("Q2") = max_percent
            Range("P2") = big_ticket
        End If
'returning min
        If Cells(r + 1, 11).Value < min_percent Then
            min_percent = Cells(r + 1, 11)
            low_ticket = Cells(r + 1, 9)
        Else
            Range("Q3") = min_percent
            Range("P3") = low_ticket
        End If
'returning max volume
        If Cells(r + 1, 12).Value > max_vol Then
            max_vol = Cells(r + 1, 12)
            vol_ticket = Cells(r + 1, 9)
        Else
            Range("Q4") = max_vol
            Range("P4") = vol_ticket
        End If
'conditional formatting
        If Cells(r, 10).Value < 0 Then
            Cells(r, 10).Interior.Color = vbRed
        ElseIf Cells(r, 10).Value > 0 Then
            Cells(r, 10).Interior.Color = vbGreen
        End If
    Next r
  

  'make numbers percents
  Range(Cells(2, 11), Cells(last_row, 11)).NumberFormat = "0.00%"
  Range("Q2:Q3").NumberFormat = "0.00%"


    
End Sub
