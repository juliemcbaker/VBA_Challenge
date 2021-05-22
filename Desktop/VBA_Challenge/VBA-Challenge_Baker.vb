Sub Pulling_Tickers()

Dim ws As Worksheet

For Each ws In Worksheets

    'initialize worksheet holder & grab worksheet name
    Dim WorkSheetName As String
    WorkSheetName = ws.Name  '**consider switching to CodeName method

    ' initialize ticker holder; **ticker is in col A/1
    Dim Current_Ticker As String
    Current_Ticker = ws.Cells(2, 1).Value

    ' initialize opening (col C/3) price holder **Needs to come from very first open;
    ' use to calculate Yearly_Change (col J/10)
    Dim Opening_Price As Double
    Opening_Price = ws.Cells(2, 3).Value

    ' initializes closing (col F/6) **Needs to come from last close of the year;
    ' use to calculate Yearly_Change (col J/10)
    ' Closing_Price will be static, but iterates through list until the end of the loop;
    Dim Closing_Price As Double

    ' initializes holder for stock volume **This will be accumulating through a loop
    ' whilst Current_Ticker & Opening_Price remain static;
    ' Comes from (col G/7); Goes into (col L/12); declaring as double fixed overflow error
    Dim Stock_Vol As Double
    Stock_Vol = 0

    ' initialize calculator for change & percentage change for the year
    ' Calculates based on Opening_Price and Closing_Price; Yearly_Change Goes into (col J/10)
    ' Percent_Change goes into(col K/11)
    Dim Yearly_Change As Double
    Dim Percent_Change As Double

    'initialize counter for last row
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' initialize location for tickers to start populating; Ticker is (col I/9)
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    ' set up summary table on the active sheet
    '*** come back later & try to format these
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Volume"

    '***JUST TEMP TO TEST LAST ROW & sheet***
    'ws.Range("H1").Value = LastRow
    'MsgBox WorkSheetName 'This is to test that it's iterating sheets

    'Loop through list (starting at 1 so the header serves to get the first instance started)
    For i = 2 To LastRow

        'check to see if we are on the same stock
        'the IF portion of the statement is what happens when the stock CHANGES
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'Before we change the ticker, we want it to
            'dump the values we have been accummulating AND need to calculate percent change
            'dump Current_Ticker; dump Openning_Price; dump Closing_Price; dump Stock_Vol
        
            'Puts ticker into summary table
            ws.Range("I" & Summary_Table_Row).Value = Current_Ticker
        
            'Calculate yearly change
            Yearly_Change = Closing_Price - Opening_Price
        
            'Puts yearly change into summary table
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                       
            'Calculate %change
            Percent_Change = (Yearly_Change / Opening_Price) * 100
        
            'Puts percent change into summary table; found formatting on stackoverflow
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
            'Puts stock volume into summary table
            ws.Range("L" & Summary_Table_Row).Value = Stock_Vol
        
            'Change the ticker AND record the opening value for the NEW stock
            ' AND update the Summary_Table_Row
            Current_Ticker = ws.Cells(i + 1, 1).Value
            Openning_Price = ws.Cells(i + 1, 3).Value
        
            Summary_Table_Row = Summary_Table_Row + 1
        
        'If the cell immediately following a row is the SAME STOCK...
        Else
            'Add to stock volume
            '***getting runtime error 6 overflow on this***
            Stock_Vol = Stock_Vol + ws.Cells(i, 7).Value
        
            'iterate through closing prices
            Closing_Price = ws.Cells(i, 6).Value
            
        End If


    Next i
    
    '_______________________________________________
    'Conditional formatting found on trumpexcel.com
    
    Dim Cll As Range
    Dim Rng As Range
    Set Rng = ws.Range("J2", ws.Range("J2").End(xlDown))
    For Each Cll In Rng
        If Cll.Value < 0 Then
            Cll.Interior.Color = vbRed
        ElseIf Cll.Value > 0 Then
            Cll.Interior.Color = vbGreen
        End If
    Next Cll
    '=================================================
    
    '====================================================
    'GENERATES BEST (& WORST) PERFORMERS
    'identifies standout performers from summary table

    Dim Great_Increase As Double
    Dim Great_Decrease As Double
    Dim Great_Volume As Double
    Dim Inc_Ticker As String
    Dim Dec_Ticker As String
    Dim Vol_Ticker As String
    Dim SummaryLastRow As Long

    'determine number of rows in summary table
    'add 1 to get first empty row
    SummaryLastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row + 1

    'MsgBox ("Start SuperSummary") 'testing during dev
    'create sub-table; Cols O/15, P/16, Q/17
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"

    'search table for largest increase
    'found code on AutomateExcel.com
    Great_Increase = Application.Max(ws.Columns("K"))
    ws.Cells(2, 17).Value = Great_Increase
    ws.Cells(2, 17).NumberFormat = "0.00%"
    
    'MsgBox Great_Increase 'just a test during development

    'search table for largest decrease
    Great_Decrease = Application.Min(ws.Columns("K"))
    ws.Cells(3, 17).Value = Great_Decrease
    ws.Cells(3, 17).NumberFormat = "0.00%"

    'search table for largest volume
    Great_Volume = Application.Max(ws.Columns("L"))
    ws.Cells(4, 17).Value = Great_Volume

    For j = 1 To SummaryLastRow

        'Check if number matches Great_Increase
        If ws.Cells(j, 11).Value = Great_Increase Then
        
            'if true, retrieve ticker
            Inc_Ticker = ws.Cells(j, 9).Value
            'MsgBox (Inc_Ticker) 'just a test during dev
            ws.Cells(2, 16).Value = Inc_Ticker
            'MsgBox ("Dude")  'just a test during dev
        
        'Check if number matches Great_Decrease
        ElseIf ws.Cells(j, 11).Value = Great_Decrease Then
            'if true, retrieve ticker
            Dec_Ticker = ws.Cells(j, 9).Value
            ws.Cells(3, 16).Value = Dec_Ticker
    
        'Check if number matches Great_Volume
        ElseIf ws.Cells(j, 12).Value = Great_Volume Then
            'if true, retrieve ticker
            Vol_Ticker = ws.Cells(j, 9).Value
            ws.Cells(4, 16).Value = Vol_Ticker
    
        'Ends If/Else conditionals
        End If
    
    Next j
    '========================================================
    '____________________________________________
    'auto adjust column width found at analysistabs.com
            
            ws.Columns("I:L").AutoFit
            ws.Columns("O:Q").AutoFit
    


'MsgBox ("Next One") 'used to make sure iterating sheets during dev


Next ws

MsgBox ("You can now view a summary of each year on its sheet.")

End Sub


