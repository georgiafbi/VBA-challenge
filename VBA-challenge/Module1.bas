Attribute VB_Name = "Module1"
Sub ClrCntnts()
'creates new sheet called "combined"
'Application.DisplayAlerts = False
'Sheets("Combined").Delete
'Sheets.Add(After:=Sheets("2014")).Name = "Combined"
'Sheets("2016").Range("A1:G1").Copy Destination:=Sheets("Combined").Range("A1:G1")

'loops through all worksheets and clears contents and resets formats
For Each ws In Worksheets

'Clears all worksheets
    'ws.Activate
    
    'selects current worksheet
    ws.Select
    
    'pauses program to switch sheets
    Application.Wait (Now + TimeValue("00:00:4"))
    
    'clears calculated data from previous program run
    ws.Range("I1:R" & ws.Cells(Rows.Count, 1).End(xlUp).Row).ClearFormats
    ws.Range("I2:R" & ws.Cells(Rows.Count, 1).End(xlUp).Row).clearContents
    
    'lables new column headers on row 1, columns I through Q
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'lables category to column O row 2
    ws.Range("O2").Value = "Greatest Percentage Increase"
    
    'lables category to column O row 4
    ws.Range("O4").Value = "Greatest Total Volume"
    
    'lables category to column O row 3
    ws.Range("O3").Value = "Greatest Percentage Decrease"
    
    ws.Range("O10").Value = "Program Run Time"
    
    'fits cells to data size
    'Sheets("Combined").Range("A1:R200").Columns.AutoFit
    
    'fits cells to data size
    ws.Range("A1:R" & ws.Cells(Rows.Count, 1).End(xlUp).Row).Columns.AutoFit
    
    'allows for animation to catch up
    Application.Wait (Now + TimeValue("00:00:02"))
    
Next ws

End Sub
Sub TckrYrlyPrcntge()

For Each ws In Worksheets
    
    Dim yrOpnVle, yrClseVle, yrlyChnge, lstRw As Double

    'ws.Activate

    'selects current worksheet
    'ws.Select

    'pauses program to switch sheets
    'Application.Wait (Now + TimeValue("00:00:05"))

    'Initializes variables to store the year open value of a stock to the year end closing value of a stock
    yrOpnVle = 0

    yrOpnVle = 0

    'Initializes a Variable to calculate the yearly change for a stock
    yrlyChnge = 0
    
    'This counter is used to track where to paste the ticker symbol of each cell in column I in the Combined sheet
    tckrCllCnt = 2

    'finds last row of combined sheet
    lstRw = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'scans through the ticker symbols in each worksheet checks to see
    'if the tickers are different and assigns the unqiue symbol to column I in the Combined worksheet
        For i = 2 To lstRw
            
            'grabs ticker symbol from 1st cell
            tckr1 = ws.Range("A" & i).Value
    
            'grabs ticker ticker symbol from 2nd cell
            tckr2 = ws.Range("A" & (i + 1)).Value
    
            'grabs first yrOpnVle for the first ticker symbol
            If i = 2 Then
                yrOpnVle = ws.Range("C2").Value
            End If
    
            'compares if the ticker symbols are different
            If tckr1 <> tckr2 Then
    
                'Grabs the close value for the tckr1
                yrClseVle = ws.Range("F" & i).Value
            
                'Calculates the yearly change per unique stock
                yrlyChnge = yrClseVle - yrOpnVle
            
                'Calculates the percentage change per unique stock
                If yrOpnVle <> 0 Then
                    prcntChnge = (yrClseVle - yrOpnVle) / yrOpnVle
                End If
    
                'Copies the unique Ticker symbols per sheet into the I column in the Combined worksheet
                ws.Range("A" & i).Copy Destination:=ws.Range("I" & tckrCllCnt)
            
                'Stores yearly change in J column of the Combined sheet
                ws.Range("J" & tckrCllCnt) = FormatCurrency(yrlyChnge)
            
                'formats cell color based on positive or negative values in yrlyChnge
                If yrlyChnge > 0 Then
                    ws.Range("J" & tckrCllCnt).Interior.ColorIndex = 10
                ElseIf yrlyChnge < 0 Then
                    ws.Range("J" & tckrCllCnt).Interior.ColorIndex = 9
                Else
                    ws.Range("J" & tckrCllCnt).Interior.ColorIndex = 15
                End If
            
                'Stores percentage change in K column of the Combined sheet
                ws.Range("K" & tckrCllCnt) = FormatPercent(prcntChnge)
            
                'Grabs the open value for the tckr2
                yrOpnVle = ws.Range("C" & (i + 1)).Value
            
                'Increments TickerCell count to track where to paste the next worksheet's ticker symbol on the Combined sheet in the next loop of the for
                tckrCllCnt = tckrCllCnt + 1
            End If
         
    Next i

    'fits cells to data size
    ws.Range("A1:R" & ws.Cells(Rows.Count, 1).End(xlUp).Row).Columns.AutoFit

    'pauses program to switch sheets
    'Application.Wait (Now + TimeValue("00:00:15"))

Next ws

End Sub

Sub TtlStckVlme2014()

Dim tckr As String
Dim ttlStckVlme As Long
Dim lstWrkshtRw, lstTckrRw As Long

Set ws = Sheets("2014")

'selects current worksheet
'ws.Select

'pauses program to switch sheets
'Application.Wait (Now + TimeValue("00:00:10"))

'Loops through each worksheet to total the stock volumne per unique ticker symbol
'For Each ws In Worksheets

'Gets the last row for ticker column I
lstTckrRw = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row

'Get the last row of the sheet Combined
lstWrkshtRw = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

For i = 2 To lstTckrRw

    'initial ttlStckVlme
    ttlStckVlme = 0
    
    'assigns value of Ticker column I to a variable tckr
    tckr = ws.Range("I" & i).Value
    
    'assigns Total Stock Volumn column L with the sum of all <vol> values for each <ticker> symbol that matches
    'the ticker symbol in the Ticker column I
    ws.Range("L" & i).Value = Application.WorksheetFunction.SumIf(Range("A2:A" & lstWrkshtRw), tckr, Range("G2:G" & lstWrkshtRw))

Next i

'fits cells to data size
ws.Range("A1:R" & ws.Cells(Rows.Count, 1).End(xlUp).Row).Columns.AutoFit

'pauses program to switch sheets
'Application.Wait (Now + TimeValue("00:00:08"))
'Next ws
End Sub
Sub TtlStckVlme2015()

Dim tckr As String
Dim ttlStckVlme As Long
Dim lstWrkshtRw, lstTckrRw As Long

Set ws = Sheets("2015")
 
'selects current worksheet
'ws.Select

'pauses program to switch sheets
'Application.Wait (Now + TimeValue("00:00:10"))

'Loops through each worksheet to total the stock volumne per unique ticker symbol
'For Each ws In Worksheets

'Gets the last row for ticker column I
lstTckrRw = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row

'Get the last row of the sheet Combined
lstWrkshtRw = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

For i = 2 To lstTckrRw

    'initial ttlStckVlme
    ttlStckVlme = 0
    
    'assigns value of Ticker column I to a variable tckr
    tckr = ws.Range("I" & i).Value
    
    'assigns Total Stock Volumn column L with the sum of all <vol> values for each <ticker> symbol that matches
    'the ticker symbol in the Ticker column I
    ws.Range("L" & i).Value = Application.WorksheetFunction.SumIf(Range("A2:A" & lstWrkshtRw), tckr, Range("G2:G" & lstWrkshtRw))

Next i

'fits cells to data size
ws.Range("A1:Q" & ws.Cells(Rows.Count, 1).End(xlUp).Row).Columns.AutoFit
'Next ws

'pauses program to switch sheets
'Application.Wait (Now + TimeValue("00:00:08"))
End Sub
Sub TtlStckVlme2016()

Dim tckr As String
Dim ttlStckVlme As Long
Dim lstWrkshtRw, lstTckrRw As Long

Set ws = Sheets("2016")
 
'ws.Activate
'selects current worksheet
'ws.Select
 
'pauses program to switch sheets
'Application.Wait (Now + TimeValue("00:00:10"))
'Loops through each worksheet to total the stock volumne per unique ticker symbol
'For Each ws In Worksheets

'Gets the last row for ticker column I
lstTckrRw = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row

'Get the last row of the sheet Combined
lstWrkshtRw = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

For i = 2 To lstTckrRw

    'initial ttlStckVlme
    ttlStckVlme = 0
    
    'assigns value of Ticker column I to a variable tckr
    tckr = ws.Range("I" & i).Value
    
    'assigns Total Stock Volumn column L with the sum of all <vol> values for each <ticker> symbol that matches
    'the ticker symbol in the Ticker column I
    ws.Range("L" & i).Value = Application.WorksheetFunction.SumIf(Range("A2:A" & lstWrkshtRw), tckr, Range("G2:G" & lstWrkshtRw))

Next i

'adjust cells to fit data size
ws.Range("A1:R" & ws.Cells(Rows.Count, 1).End(xlUp).Row).Columns.AutoFit
'Next ws
'pauses program for animation to catch up
'MsgBox (Now + TimeValue("00:00:10"))
'Application.Wait (Now + TimeValue("00:00:7"))
End Sub

Sub GrtstPrcntgeIncrse()

Dim lstRw As Long
Dim mxVle, crrntCll, gPI As Double
Dim mxVleTckr, gPIYr, gPITckr As String

gPDI = 0

For Each ws In Worksheets

'selects current worksheet
'ws.Select

'pauses program to switch sheets
'Application.Wait (Now + TimeValue("00:00:3"))


'initializes the variable mxVle
mxVle = 0

'assigns the initial <ticker> cell value to mxVleTckr
mxVleTckr = "A2"

'Sets the Combined worksheet as sht
'Set sht = Sheets("Combined")

'finds the last row in column K
lstRw = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row

'loops through each cell in column Percentage Change (K)
For i = 2 To lstRw
    
    'stores current cell value into crrntCll variable
    crrntCll = ws.Range("K" & i).Value
    'if current cell value is greater than last value store it in mxVle
    If crrntCll > mxVle Then
    
        mxVle = crrntCll
        mxVleTckr = "I" & i
        
    End If
Next i

'Finds the greatest percentage Increase in all worksheets
If gPI < mxVle Then
    gPI = mxVle
    gPIYr = ws.Name
    gPITckr = ws.Range(mxVleTckr).Value
End If

'assigns minimum value and ticker symbols to Greatest % Increase section
ws.Range("Q2").Value = FormatPercent(mxVle)
ws.Range("P2").Value = ws.Range(mxVleTckr).Value

'fits cells to data size
ws.Range("A1:R" & ws.Cells(Rows.Count, 1).End(xlUp).Row).Columns.AutoFit

'pauses program for animation to catch up
'Application.Wait (Now + TimeValue("00:00:2"))
Next ws

''assigns greatest percentage increase and sheet name to the Combined worksheet column
'Sheets("Combined").Range("P2").Value = gPITckr
'Sheets("Combined").Range("Q2").Value = FormatPercent(gPI)
'Sheets("Combined").Range("R2").Value = gPIYr
''fits cells to data size
'Sheets("Combined").Range("A1:R30000").Columns.AutoFit
End Sub


Sub GrtstPrcntgeDcrse()

'Dim lstRw As Long
'Dim mnVle, crrntCll, gPD As Double
Dim mnVleTckr, gPDYr, gPDTckr As String

'gPD = 0

For Each ws In Worksheets

'selects current worksheet
'ws.Select

'pauses program to switch sheets
'Application.Wait (Now + TimeValue("00:00:03"))

'initializes the variable mnVle
mnVle = ws.Range("K2").Value

'assigns the initial <ticker> cell value to mnVleTckr
mnVleTckr = "A2"

'Sets the Combined worksheet as sht
'Set sht = Sheets("Combined")

'finds the last row in column K
lstRw = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row

'loops through each cell in column Percentage Change (K)
For i = 2 To lstRw
    
    'stores current cell value into crrntCll variable
    crrntCll = ws.Range("K" & i).Value
    
    'if current cell value is smaller than last value store it in mnVle
    If crrntCll < mnVle Then
    
        mnVle = crrntCll
        mnVleTckr = "I" & i
        
    End If
Next i

'assigns minimum value and ticker symbols to Greatest % Decrease section
ws.Range("Q3").Value = FormatPercent(mnVle)

'Finds the greatest percentage decrease in all worksheets
'If gPD > mnVle Then
'    gPD = mnVle
'    gPDYr = ws.Name
'    gPDTckr = ws.Range(mnVleTckr).Value
'End If

ws.Range("P3").Value = ws.Range(mnVleTckr).Value

'fits cells to data size
ws.Range("A1:R" & ws.Cells(Rows.Count, 1).End(xlUp).Row).Columns.AutoFit

'pauses program for animation to catchup
'Application.Wait (Now + TimeValue("00:00:2"))

Next ws

'assigns greatest percentage decrease and sheet name to the Combined worksheet column
'Sheets("Combined").Range("P3").Value = gPDTckr
'Sheets("Combined").Range("Q3").Value = FormatPercent(gPD)
'Sheets("Combined").Range("R3").Value = gPDYr
'
''fits cells to data size
'Sheets("Combined").Range("A1:R30000").Columns.AutoFit
End Sub

Sub GrtstTtlVlme()

Dim mxTtlVleTckr, mxGTVYr, mxGTVTckr As String

mxGTV = 0

For Each ws In Worksheets

    'selects current worksheet
    'ws.Select

    'pauses program to switch sheets
    'Application.Wait (Now + TimeValue("00:00:03"))

    'initializes the variable mnVle
    mxTtlVle = 0

    'Sets the Combined worksheet as sht
    'Set sht = Sheets("Combined")

    'finds the last row in column K
    lstRw = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row

    'loops through each cell in column Percentage Change (K)
    For i = 2 To lstRw
    
        'stores current cell value into crrntCll variable
        crrntCll = ws.Range("L" & i).Value
    
        'if current cell value is smaller than last value store it in mnVle
        If crrntCll > mxTtlVle Then
    
            mxTtlVle = crrntCll
            mxTtlVleTckr = "I" & i
        
        End If
    Next i

    'assigns minimum value and ticker symbols to Greatest % Decrease section
    ws.Range("Q4").Value = mxTtlVle
    
    'finds the max total stock volume of all sheets stores in a variable called mxGTV
    'along with the name of the sheet and ticker symbol
    If mxGTV < mxTtlVle Then
        mxGTV = mxTtlVle
        mxGTVYr = ws.Name
        mxGTVTckr = ws.Range(mxTtlVleTckr).Value
    End If
    
    ws.Range("P4").Value = ws.Range(mxTtlVleTckr).Value

    'fits cells to data
    ws.Range("A1:R" & ws.Cells(Rows.Count, 1).End(xlUp).Row).Columns.AutoFit

    'pauses program for sheet animation to catch up
    'Application.Wait (Now + TimeValue("00:00:2"))

Next ws

'adds the greatest total stock volume to the "Combined" sheet with ticker and year
'Sheets("Combined").Range("P4").Value = mxGTVTckr
'Sheets("Combined").Range("Q4").Value = mxGTV
'Sheets("Combined").Range("R4").Value = mxGTVYr

'fits cells to data size
'Sheets("Combined").Range("A1:R30000").Columns.AutoFit
End Sub

Sub StckMrktPrgrm()

Dim strtTme As Double
Dim mntsElpsd As String

'used to track the time it takes for this program to run
strtTme = Timer
MsgBox ("This program may take between 8 to 13 minutes to run. Data will be reset first. Click OK to continue.")
'calls macros to run full program
Call ClrCntnts 'Macro1
Sheets("2016").Select
'pauses program for sheet animation to catch up
Application.Wait (Now + TimeValue("00:00:2"))
Call TckrYrlyPrcntge 'Macro2
Call TtlStckVlme2016 'Macro (1/3)*3
Call TtlStckVlme2015 'Macro (2/3)*3
Call TtlStckVlme2014 'Macro (3/3)*3
Call GrtstPrcntgeIncrse 'Macro4
Call GrtstPrcntgeDcrse 'Macro5
Call GrtstTtlVlme 'Macro6

'calculates the total time elapsed
mntsElpsd = Format((Timer - strtTme) / 86400, "hh:mm:ss")
'stores program run time on each sheet
For Each ws In Worksheets
    ws.Select
    
    'pauses program for sheet animation to catch up
    Application.Wait (Now + TimeValue("00:00:2"))
    
    ws.Range("P10").Value = mntsElpsd
    ws.Range("P9").Value = "Time (Minutes)"
    'pauses program for sheet animation to catch up
    'Application.Wait (Now + TimeValue("00:00:2"))
    
    'fits cells to data
    ws.Range("A1:R" & ws.Cells(Rows.Count, 1).End(xlUp).Row).Columns.AutoFit
Next ws
'displays the time to the user
MsgBox "This program took " & mntsElpsd & " minutes", vbInformation
End Sub

