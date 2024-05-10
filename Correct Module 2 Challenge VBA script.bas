Attribute VB_Name = "Module1"
Sub MultipleYearStockData()

' LOOP THROUGH ALL SHEETS
' --------------------------------------------
For k = 1 To 4
'k = 1
Dim WorksheetName As String
' Determine the Last Row

'lastrow = Worksheets(k).Cells(Worksheets(k).Count, 1).End(xlUp).Row
' Grabbed the WorksheetName
lastrow = 93001
'Label Top Columns
Worksheets(k).Range("I1").Value = "Ticker"
Worksheets(k).Range("J1").Value = "Quarterly Change"
Worksheets(k).Range("K1").Value = "Percentage Change"
Worksheets(k).Range("L1").Value = "Total Stock Volume"
Worksheets(k).Range("P1").Value = "Ticker"
Worksheets(k).Range("Q1").Value = "Value"
Worksheets(k).Cells(2, 15).Value = "Greatest % Increase"
Worksheets(k).Cells(3, 15).Value = "Greatest % Decrease"
Worksheets(k).Cells(4, 15).Value = "Greatest Total Volume"
    
    'Set "Ticker" column as letters
    Dim Ticker As String
    Ticker = " "
    
    'Set "TicketRow" as starting row #
    Dim TickerRow As Long
    TickerRow = 1
    
    'Set "Opening" & "Closing" as numbers that can be subtracted ("long" doesn't work)
    Dim Opening As Double
    Dim Closing As Double
    
    'Set "StartingPrice" as starting row #
    Dim StartingPrice As Long
    StartingPrice = 2
    
    'Set "Quarterly_Change" as number for the solution of the difference between Closing - Starting
    Dim Quarterly_Change As Double
    'Set "Quarterly_Change_Row" as starting row #
    Dim Quarterly_Change_Row As Long
    Quarterly_Change_Row = 0
    
    Dim Percentage_Change As Double
    Dim Percentage_Change_Row As Long
    Percentage_Change_Row = 0
    
    Dim Volume_Amount As Double
    Dim VolumeRow As Long
    VolumeRow = 2
    
    Dim StartingVolume As Long
    StartingVolume = 2
    
    Dim Total_Stock_Volume As Double
    Dim Total_Stock_Volume_Row As Long
    TickerRow = 2
    Total_Stock_Volume = 0
    For i = 2 To lastrow
                          
            If Worksheets(k).Cells(i + 1, 1).Value <> Worksheets(k).Cells(i, 1).Value Then
            
            Ticker = Worksheets(k).Cells(i, 1).Value
            Worksheets(k).Cells(TickerRow, 9).Value = Ticker
                
            Opening = Worksheets(k).Cells(StartingPrice, 3).Value
            StartingPrice = i + 1
            Closing = Worksheets(k).Cells(i, 6).Value
            
            Quarterly_Change = Closing - Opening
            Quarterly_Change_Row = Quarterly_Change_Row + 1
            Worksheets(k).Cells(TickerRow, 10).Value = Quarterly_Change
            
            Percentage_Change = (Closing - Opening) / Opening * 100
            Worksheets(k).Cells(TickerRow, 11).Value = Percentage_Change

            Total_Stock_Volume = Total_Stock_Volume + Worksheets(k).Cells(i, 7).Value
            Worksheets(k).Cells(TickerRow, 12).Value = Total_Stock_Volume
             Total_Stock_Volume = 0
            TickerRow = TickerRow + 1
             Else
             Total_Stock_Volume = Total_Stock_Volume + Worksheets(k).Cells(i, 7).Value

        If Worksheets(k).Cells(i, 10).Value > 0 Then
        Worksheets(k).Cells(i, 10).Interior.ColorIndex = 4
        ElseIf Worksheets(k).Cells(i, 10).Value < 0 Then
        Worksheets(k).Cells(i, 10).Interior.ColorIndex = 3
        End If
       
       End If

    Next i
            greatest_increase = -1
            For i = 2 To 1501
            If Worksheets(k).Cells(i, 11).Value > greatest_increase Then
            greatest_ticker = Worksheets(k).Cells(i, 9).Value
            greatest_increase = Worksheets(k).Cells(i, 11).Value
            End If
            Next i
            Worksheets(k).Cells(2, 17).Value = greatest_increase
            Worksheets(k).Cells(2, 16).Value = greatest_ticker
              
             greatest_decrease = -1
            For i = 2 To 1501
            If Worksheets(k).Cells(i, 11).Value < greatest_decrease Then
               greatest_ticker = Worksheets(k).Cells(i, 9).Value
               greatest_decrease = Worksheets(k).Cells(i, 11).Value
            End If
            Next i
            Worksheets(k).Cells(3, 17).Value = greatest_decrease
            Worksheets(k).Cells(3, 16).Value = greatest_ticker
            
            greatest_increase = -1
            For i = 2 To 1501
            If Worksheets(k).Cells(i, 12).Value > greatest_increase Then
               greatest_ticker = Worksheets(k).Cells(i, 9).Value
               greatest_increase = Worksheets(k).Cells(i, 12).Value
            End If
            Next i
            Worksheets(k).Cells(4, 17).Value = greatest_increase
            Worksheets(k).Cells(4, 16).Value = greatest_ticker
              
Next k
          

End Sub
