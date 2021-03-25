Attribute VB_Name = "Module1"
Sub UpdateAllSheets()

'CalculateRunTime

    Dim StartTime As Double
    Dim MinutesElapsed As String
    
'Remember time when macro starts
    StartTime = Timer

'Update all worksheets

    Dim ws As Worksheet
    Application.ScreenUpdating = False
    For Each ws In Worksheets
        ws.Select
        Call Stocks
    Next
    Application.ScreenUpdating = True
    
'Determine how many seconds code took to run
    MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")

'Notify user in seconds
    MsgBox "This code ran successfully in " & MinutesElapsed & " minutes", vbInformation

End Sub

Sub Stocks()

'Declare Varialbles
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim Volume As LongLong
    Dim SummaryTableRow As Integer
    Dim LastRow As Long
    Dim YearOpen As Double
    Dim YearClosed As Double

' Set Variables

    Volume = 0
    x = 2
     
'Add Summary Table Column Headers
    SummaryTableRow = 2
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
     
    ' Loop through table
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To LastRow
        
        ' Check if still within the same ticker, if not then...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                YearOpen = Cells(x, 3).Value
                YearClosed = Cells(i, 6).Value
    
        ' Add Ticker values to Summary Table
                Ticker = Cells(i, 1).Value
                Cells(SummaryTableRow, 9).Value = Ticker
            
        ' Add YearlyChange value to Summary Table = Total Close - Total Open
                YearlyChange = YearClosed - YearOpen
                Cells(SummaryTableRow, 10).Value = YearlyChange
            
        ' Add PercentChange value to Summary Table = YearClosed - YearOpen / YearOpen *100
                If YearOpen = 0 Then
                    PercentChange = 0
                Else
                    PercentChange = YearlyChange / YearOpen
                    Cells(SummaryTableRow, 11).Value = PercentChange
                    Columns("K").NumberFormat = "0.00%"
                End If
            
        ' Add Volume value to Summary Table = Total Open + Total Close
                Volume = Volume + Cells(i, 7).Value
                Cells(SummaryTableRow, 12).Value = Volume
    
                SummaryTableRow = SummaryTableRow + 1
                x = (i + 1)
        
        ' Otherwise continue
            Else
                Volume = Volume + Cells(i, 7).Value
   
        End If

        ' Finish Loop
        Next i
         

 'Conditional formatting that will highlight positive change in green and negative change in red
    
    LastRow = Cells(Rows.Count, "J").End(xlUp).Row
    
        For b = 2 To LastRow
    
            If Cells(b, 10).Value >= 0 Then
                Cells(b, 10).Interior.ColorIndex = 4
            Else
                Cells(b, 10).Interior.ColorIndex = 3
        End If
        
     Next b
        
        
 'Challenge
 
    ' Loop through table
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
    'Add Summary Table Row Headers
        Cells(2, 14).Value = "Greatest % Increase"
        Cells(3, 14).Value = "Greatest % Decrease"
        Cells(4, 14).Value = "Greatest Total Volume"
        
    'Add Summary Table Column Headers
        Cells(1, 15).Value = "Ticker"
        Cells(1, 16).Value = "Value"
    
    'Return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume"
        Cells(2, 16).Value = WorksheetFunction.Max(Range("K:K"))
        Cells(3, 16).Value = WorksheetFunction.Min(Range("K:K"))
        Cells(4, 16).Value = WorksheetFunction.Max(Range("L:L"))
     
        ' Loop
            For c = 2 To LastRow
     
                If Cells(c, 11).Value = Cells(2, 16).Value Then
                    Cells(2, 15).Value = Cells(c, 9).Value
                ElseIf Cells(c, 11).Value = Cells(3, 16).Value Then
                    Cells(3, 15).Value = Cells(c, 9).Value
                ElseIf Cells(c, 12).Value = Cells(4, 16).Value Then
                    Cells(4, 15).Value = Cells(c, 9).Value
            End If
        
         Next c
     
     'Format Columns
        Cells(2, 16).NumberFormat = "0.00%"
        Cells(3, 16).NumberFormat = "0.00%"
        Cells(4, 16).NumberFormat = "0.00"
     
    
End Sub
