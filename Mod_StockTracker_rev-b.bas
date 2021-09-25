Attribute VB_Name = "Mod_StockTracker"
Sub StockTracker()

    Dim Tic, TicInc, TicDec, TicVol As String
    Dim Ope, Clo, Vol, Chg, pChg, Inc_pChg, Dec_pChg, tVol, StartTime, EndTime As Double
    Dim LRow, iRow, jRow, LSheet, iSheet As Integer
    
    'store starting time
    StartTime = Timer
    Application.ScreenUpdating = False
     
    'determine last sheet
    LSheet = Sheets.Count
    
    'set intial values for greatest values *bonus*
    Inc_pChg = 0
    Dec_pChg = 0
    tVol = 0
    
    'iterate through each sheet
    For iSheet = 1 To LSheet
        
        With Sheets(iSheet)
        'set initial values
            LRow = .Cells(1, 1).End(xlDown).Row
            Vol = 0
            jRow = 2
            
            'header formatting
            .Cells(1, 9).Value = "Ticker"
            .Cells(1, 10).Value = "Yearly Change"
            .Cells(1, 11).Value = "% Change"
            .Cells(1, 12).Value = "Total Stock Volume"
            
            'for each row add the openeing, closing, and volume.
            For iRow = 2 To LRow
                
                'status tracker so you know Excel is actually doing something and not just frozen
                Application.StatusBar = "Calculating... " & Round((iRow / LRow) * 100, 0) & "% completed on sheet " & iSheet & " of " & LSheet
                
                Tic = .Cells(iRow, 1).Value
                Vol = Vol + .Cells(iRow, 7).Value
                
                'Store 1st opening value per ticker symbol
                If .Cells(iRow, 1).Value <> .Cells(iRow - 1, 1).Value Then
                    Ope = .Cells(iRow, 3).Value
                
                'If the next ticker symbol is different calculate the difference and % difference for opening and closing and print the values to cells
                ElseIf .Cells(iRow, 1).Value <> .Cells(iRow + 1, 1).Value Then
                    
                    Clo = .Cells(iRow, 6).Value
                    Chg = Clo - Ope
                    
                    'error handler in case opening price is 0
                    If Ope <> 0 Then
                        pChg = (Clo - Ope) / Ope
                    Else
                        pChg = 1
                    End If
                    
                    'print values to cells
                    .Cells(jRow, 9).Value = Tic
                    .Cells(jRow, 11).Value = FormatPercent(pChg)
                    .Cells(jRow, 12).Value = Vol
                    
                    With .Cells(jRow, 10)
                        .Value = Chg
                        'color formatting for yearly change
                        If Chg < 0 Then
                            .Interior.Color = vbRed
                        Else
                            .Interior.Color = vbGreen
                        End If
                    End With
                    
                    'checks for greatest increase / decrease / volume *bonus*
                    If pChg > Inc_pChg Then
                        TicInc = Tic
                        Inc_pChg = pChg
                    ElseIf pChg < Dec_pChg Then
                        TicDec = Tic
                        Dec_pChg = pChg
                    ElseIf Vol > tVol Then
                        TicVol = Tic
                        tVol = Vol
                    End If
                                
                    'add to summary counter and clear variables
                    jRow = jRow + 1
                    Ope = 0
                    Clo = 0
                    Vol = 0
                    
                End If
                
            Next iRow
        
        .Range("A:P").Columns.AutoFit
        
        End With
    
    Next iSheet
    
    'print greatest values *bonus*
    With Sheets(1)
        .Cells(1, 15).Value = "Ticker"
        .Cells(1, 16).Value = "Value"
        .Cells(2, 14).Value = "Greatest % Increase"
        .Cells(2, 15).Value = TicInc
        .Cells(2, 16).Value = FormatPercent(Inc_pChg)
        .Cells(3, 14).Value = "Greatest % Decrease"
        .Cells(3, 15).Value = TicDec
        .Cells(3, 16).Value = FormatPercent(Dec_pChg)
        .Cells(4, 14).Value = "Greatest Total Volume"
        .Cells(4, 15).Value = TicVol
        .Cells(4, 16).Value = tVol
        .Range("A:P").Columns.AutoFit
    End With

    Application.StatusBar = False
    EndTime = Round(Timer - StartTime, 2)
    MsgBox ("Completed in " & EndTime & " seconds...")
    Application.ScreenUpdating = True

End Sub
