Attribute VB_Name = "Module1"
Sub pushBackTimeForAllTweetsNYTime()
LastRow = Range("A" & Rows.Count).End(xlUp).Row

Range("D4").Resize(LastRow - 3).Formula = "=IF(A4=A1, A4, A4-1)"
Range("D1").Value = Range("A1").Value
Range("D2").Value = Range("A2").Value
Range("D3").Value = Range("A3").Value

End Sub

Sub cleanLastSpaceFromColumnC()
Dim MyRange As Range
Dim lngRow As Long

    Dim cl As Range
    Set MyRange = Range("C1", "C" & Range("C" & Rows.Count).End(xlUp).Row)
    MyRange.Select
    
    For Each cl In Selection
        cl.Activate
        If ActiveCell.Value = " " Then ActiveCell.Value = ""

    Next cl
        Range("A1").Select


End Sub

Sub fillInDateInColumnA_AfterSvetWithCurrentDateNYC()
'Note, this could be dangerous to use the location is not New York NYSE (for example Japan)



LastRow = Range("A" & Rows.Count).End(xlUp).Row
Cells(LastRow + 1, "A").Value = Format(Now, "yyyy/mm/dd")

        Range("A1").Select


End Sub
Sub copySheet()
'
' copySheet Macro
'

'
    ActiveSheet.Select
    
    ActiveSheet.Copy After:=Sheets(6)
    Application.GoTo Reference:="copySheet"
    Sheets("perStockTweets").Select
    ActiveWindow.ScrollWorkbookTabs Position:=xlLast
    
End Sub




Sub CSV_Import_Loop()

    Dim sh As Worksheet, sPath As String, sName As String
    Dim ws As Worksheet
    Dim cumFile As String
    Dim csvFileName As String
    Dim csvFileNameNoExtName As String
    Dim CSVs()  As String, i As Long
    Dim CSVNameOnlys()  As String
    
    Set sh = ActiveWorkbook.Sheets("perStockTweets") 'set to current worksheet name
    sPath = "C:\Users\Bill\Desktop\longTermTest\"
    sName = Dir(sPath & "*.csv")
    
    '// Store each CSV file in an array
    Do While sName <> ""
        cumFile = sPath & sName
        i = i + 1
        ReDim Preserve CSVs(1 To i)
        ReDim Preserve CSVNameOnlys(1 To i)
        CSVs(i) = cumFile
      
        CSVNameOnlys(i) = sName
        sName = Dir()
    Loop
    
    'MsgBox ("Product position matrix: " & vbCrLf & vbCrLf & Join(CSVNameOnlys, vbCrLf))
    
    If i Then
        'loop the CSVs array
        For i = 1 To UBound(CSVs)
            cumFile = CSVs(i)
            csvFileName = CSVNameOnlys(i)
            Workbooks.Add Template:= _
                  "C:\Users\Bill\AppData\Roaming\Microsoft\Templates\sizesWithFormAndDirectoryLoopv5.xltm" 'for each file in the directory, I want to open it in a new instance of the template.
            Set sh = ActiveWorkbook.Sheets("perStockTweets")
            
            
            With sh.QueryTables.Add(Connection:="TEXT;" & cumFile, Destination:=sh.Range("A1"))
               .TextFileParseType = xlDelimited
               .TextFileCommaDelimiter = True
               .Refresh
            End With
            
            'get rid of rows that don't have a date in column B.
            LastRow = Range("A" & Rows.Count).End(xlUp).Row
            For j = 2 To LastRow
            If Not IsDate(Cells(j, "B")) Then Cells(j, 2).EntireRow.Delete
            Next
            
            'The rest comes for the "DO ALL" button on the form (except for the last line which says sName=Dir()
            'Dim ws As Worksheet
            Set ws = Worksheets("GoogleData")
            csvFileNameNoExtName = Replace(csvFileName, ".csv", "")
            'MsgBox (csvFileName)
            
            'MsgBox ("need to set ws.Cells(4,2) etc some way in tweetsLoop")
            ws.Cells(4, 2) = csvFileNameNoExtName
            ws.Cells(3, 2) = Format(Now, "yyyy/mm/dd")
            ws.Cells(2, 2) = "2009-07-15"
            'If TextBox3.Value = "" Then TextBox3.Text = Format(Now, "yyyy/mm/dd")
            
            Application.Run "generateDatesAndMbFormbPerFileFromPerl"
            Application.Run "Macro9"
            Application.Run "fromSvetToCleanUP2"
            Application.Run "saveToCSV"
        Next
    End If

End Sub


Sub remove_zero_col_D()

Dim myloop

For myloop = Range("D655536").End(xlUp).Row - 1 To 1 Step -1
    If Cells(myloop, 8).Value = 0 Then Rows(myloop).EntireRow.Delete

Next myloop


End Sub
 Sub saveToCSV()
 Dim stockName As String
 Application.DisplayAlerts = False
 stockName = ActiveWorkbook.Worksheets("GoogleData").Cells(4, 2).Text
 Sheets("sheetForSvet").Select
Sheets("sheetForSvet").Copy

ActiveWorkbook.SaveAs Filename:= _
"C:\Users\Bill\Documents\inputTestDataForModel\" & stockName & "TEST" & ".csv", FileFormat:= _
xlCSV, CreateBackup:=False
ActiveWorkbook.Close True
Application.DisplayAlerts = True
 End Sub
 
Sub CSV_Import_Tweets()
Dim ws As Worksheet, strFile As String
 
Set ws = ActiveWorkbook.Sheets("perStockTweets") 'set to current worksheet name
 ChDir "C:\Users\Bill\Desktop\longTermTest"
strFile = Application.GetOpenFilename("Text Files (*.csv),*.csv", , "Please selec text file...")
 With ws.QueryTables.Add(Connection:="TEXT;" & strFile, Destination:=ws.Range("A1"))
     .TextFileParseType = xlDelimited
     .TextFileCommaDelimiter = True
     .Refresh
 End With
   'ws.Name = "testing"
   ActiveWorkbook.Close True
   
 End Sub
Sub generateDatesAndMbFormbPerFileFromPerl()
    
    Dim objFSO      As Object
    Dim MyFolder    As String
    Dim fName       As String
    Dim i           As Long
    Dim Fyle        As Object
    Dim f()
    
    'MyFolder = "C:\Users\Bob\Desktop\Twb"            '<< adjust to suit
    MyFolder = UserForm1.TextBox5.Text
    
    If Right$(MyFolder, 1) <> "\" Then MyFolder = MyFolder & "\"
    
    If Len(Dir(MyFolder, vbDirectory)) Then
        fName = Dir(MyFolder & "*.csv")
        If Len(fName) Then
            
            Set objFSO = CreateObject("scripting.filesystemobject")
            ReDim f(1 To objFSO.getfolder(MyFolder).Files.Count, 1 To 2)
            
            For Each Fyle In objFSO.getfolder(MyFolder).Files
                i = i + 1
                f(i, 1) = Mid(Fyle.Name, 8, 8)
                f(i, 2) = Fyle.Size / 1024
            Next
        End If
        
        If i Then
            Worksheets("mbPerFileFromPerl").Range("a1:b1") = [{"Dates", "Mb"}]
            Worksheets("mbPerFileFromPerl").Range("a2").Resize(i, 2) = f
            With Worksheets("mbPerFileFromPerl").Range("a2").Resize(i)
                .TextToColumns Destination:=.Cells(1), DataType:=1, FieldInfo:=Array(1, 4)
                .NumberFormat = "mm-dd-yyyy"
            End With
        End If
    End If

End Sub

Sub OldVersiongenerateDatesAndMbFormbPerFileFromPerl()
    
    Dim objFSO      As Object
    Dim MyFolder    As String
    Dim fName       As String
   Dim dStartTime As Date
    Dim i           As Long
    Dim Fyle, f()
    
    'MyFolder = "C:\Users\Bill\Desktop\Tweets to Aug 19"            '<< adjust to suit
    MyFolder = UserForm1.TextBox5.Text
    
    If Right$(MyFolder, 1) <> "\" Then MyFolder = MyFolder & "\"
    
    If Len(Dir(MyFolder, vbDirectory)) Then
        fName = Dir(MyFolder & "*.csv")
        If Len(fName) Then
            
            Set objFSO = CreateObject("scripting.filesystemobject")
            ReDim f(1 To objFSO.getfolder(MyFolder).Files.Count, 1 To 2)
            
            For Each Fyle In objFSO.getfolder(MyFolder).Files
                i = i + 1
                
                dStartTime = DateSerial("20" & Mid(Fyle.Name, 14, 2), Mid(Fyle.Name, 11, 2), Mid(Fyle.Name, 8, 2))
                
                f(i, 1) = Format(dStartTime, "mm/dd/yyyy")
                f(i, 2) = Fyle.Size / 1024
                
            Next
            
        End If
        
        If i Then
            Worksheets("mbPerFileFromPerl").Range("a1:b1") = [{"Dates", "Mb"}]
            Worksheets("mbPerFileFromPerl").Range("a2").Resize(i, 2) = f
        End If
    End If

End Sub
 
 
 Sub CSV_Import_Sizes()
Dim ws As Worksheet, strFile As String
 
Set ws = ActiveWorkbook.Sheets("mbPerFileFromPerl") 'set to current worksheet name
    ws.Cells.ClearContents
strFile = Application.GetOpenFilename("Text Files (*.csv),*.csv", , "Please selec text file...")
 With ws.QueryTables.Add(Connection:="TEXT;" & strFile, Destination:=ws.Range("A2"))
     .TextFileParseType = xlDelimited
     .TextFileCommaDelimiter = True
     .Refresh
 End With
   'ws.Name = "testing"
   'dates   mb  countAll    adjDate
   With ws
.Cells(1, 1) = "dates"
.Cells(1, 2) = "mb"
.Cells(1, 3) = "countAll"
.Cells(1, 4) = "adjDate"

End With
 End Sub
 Sub ShowForm()
    UserForm1.Show
End Sub
 
Sub CleanUp2ForTrainingSet()
'the reason this is for the training set is because it does not leave the last row of data
'which would be the current day's data and would be needed for the test set.
 Range("C1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlToLeft
  
    Range("AH1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A2:CO22").Select
    Range("A22").Activate
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
  Rows(ActiveCell.Row).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    Range("F1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    ActiveWindow.LargeScroll ToRight:=5
    Range("CP1").Select
    ActiveSheet.Paste
    ActiveWindow.LargeScroll ToRight:=-5
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
       
End Sub
Sub fromMbToNumTweets()
LastRow = Range("B" & Rows.Count).End(xlUp).Row

Range("C2").Resize(LastRow - 1).Formula = "=7.1612*B2"

End Sub
Sub normalizeCountsWithCountAll()
Dim MyRange As Range

LastRow = Range("B" & Rows.Count).End(xlUp).Row

Range("G2").Resize(LastRow - 1).Formula = "=B2*1000000/VLOOKUP(A2, $D$2:$E$837,2,0)"
'Punt for values that have "N/A"
Set MyRange = Range("G1", "G" & Range("G" & Rows.Count).End(xlUp).Row)
For Each c In MyRange.Cells
'If c.Text = "#N/A" Then c.Value = Range("B" & c.Row, "B" & c.Row).Value * 1000000 / Range("E" & c.Row, "E" & c.Row).Value
If c.Text = "#N/A" Then c.Value = Range("B" & c.Row, "B" & c.Row).Value * 5

Next


End Sub
Sub fromCountToCountNorm()
LastRow = Range("B" & Rows.Count).End(xlUp).Row

Range("G2").Resize(LastRow - 1).Formula = "=B2*1000000/VLOOKUP(A2, $D$2:$E$837,2,0)"

End Sub
Sub deleteNonTweetRecordsFromMbPerFileFromPerl()
    
    Dim cl As Range
    
    For Each cl In Selection
        cl.Activate
        Do While ActiveCell.Value <> "" And Not ActiveCell.Value < 69000
            ActiveCell.EntireRow.Delete
        Loop
    Next cl
    
End Sub

Sub GetData()
'   thanks to Ron McEwan :^)

    Dim QuerySheet As Worksheet
    Dim DataSheet As Worksheet
    Dim EndDate As Date
    Dim StartDate As Date
    Dim Symbol As String
    Dim qurl As String
    Dim nQuery As Name
    Dim LastRow As Long
    Dim x As Range
    Dim B4 As String
    Dim i As Long, endRow As Long, j As Integer
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    complete = False
    bSymbolNotFound = False 'Greg Lovern
    
    Set DataSheet = ActiveSheet
  
        StartDate = DataSheet.Range("B2").Value
        EndDate = DataSheet.Range("B3").Value
        Symbol = DataSheet.Range("B4").Value
        Range("C7").CurrentRegion.ClearContents
        
'construct the URL for the query
        
        'Google
        qurl = "http://finance.google.com/finance/historical?q=" & Symbol
        qurl = qurl & "&startdate=" & MonthName(Month(StartDate), True) & _
               "+" & Day(StartDate) & "+" & Year(StartDate) & _
               "&enddate=" & MonthName(Month(EndDate), True) & _
               "+" & Day(EndDate) & "+" & Year(EndDate) & "&output=csv"

        'Yahoo
'        qurl = "http://ichart.finance.yahoo.com/table.csv?s=" & Symbol
'        qurl = qurl & "&a=" & Month(StartDate) - 1 & "&b=" & Day(StartDate) & _
'            "&c=" & Year(StartDate) & "&d=" & Month(EndDate) - 1 & "&e=" & _
'            Day(EndDate) & "&f=" & Year(EndDate) & "&g=" & Range("P2") & "&q=q&y=0&z=" & _
'            Symbol & "&x=.csv"
        Range("b5") = qurl
                   
QueryQuote:

            'Web query
            With ActiveSheet.QueryTables.Add(Connection:="URL;" & qurl, Destination:=DataSheet.Range("C7"))
                .BackgroundQuery = True
                .TablesOnlyFromHTML = False
                On Error GoTo BadSymbol 'Greg Lovern
                Application.Wait Now() + TimeValue("00:00:25")
                .Refresh BackgroundQuery:=False
                On Error GoTo 0 'Greg Lovern
                .SaveData = True
            End With
    
            Range("C7").CurrentRegion.TextToColumns Destination:=Range("C7"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                Semicolon:=False, Comma:=True, Space:=False, other:=False
            
            Range(Range("C7"), Range("C7").End(xlDown)).NumberFormat = "mmm d/yy"
            Range(Range("D7"), Range("G7").End(xlDown)).NumberFormat = "0.00"
            Range(Range("H7"), Range("H7").End(xlDown)).NumberFormat = "0,000"
            Range(Range("I7"), Range("I7").End(xlDown)).NumberFormat = "0.00"
            
            'If Google doesn't return "Adjusted Close", fill col I with "Close" values
            endRow = Range("G65536").End(xlUp).Row
            If DataSheet.Cells(endRow, "I") = "" Then
               For i = 7 To endRow
                   DataSheet.Cells(i, "I").Value = DataSheet.Cells(i, "G").Value
               Next
            End If


    With ThisWorkbook
        For Each nQuery In Names
            If IsNumeric(Right(nQuery.Name, 1)) Then
                nQuery.Delete
            End If
        Next nQuery
    End With
    
'turn calculation back on
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    
'    Range("C7:I2000").Select
'    Selection.Sort Key1:=Range("C8"), Order1:=xlAscending, Header:=xlGuess, _
'        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    Range("C7:I2000").Sort Key1:=Range("C8"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom 'Greg Lovern

'    Range("C1").Select
'    Selection.ColumnWidth = 12
    Range("C1").ColumnWidth = 12 'Greg Lovern
    
    'UpdateScale
    'UpdateScale2
    'UpdateScale3
    Range("B4").Select
  
LastRow = Cells(Rows.Count, "I").End(xlUp).Row
Range("BG7").FormulaR1C1 = "=AVERAGE(R" & LastRow - Range("P5") + 1 & "C[-50]:R" & LastRow & "C[-50])"

'On Error Resume Next
'Range("H4").ClearContents
'Set x = Range("I" & Rows.Count).End(xlUp)
'Range("H4") = x / x.Offset(-Range("L6").Value)
'On Error Resume Next

'If Sheets("Candles").Range("B4").Value = "DIA" Then
'
'    Sheets("Candles").Range("F4").ClearContents 'entry date dow return.
'    Set X = Sheets("Candles").Range("I" & Rows.Count).End(xlUp)
'
'    Sheets("Candles").Range("F4") = X / X.Offset(-Sheets("Candles").Range("L6").Value) - 1
'    Sheets("Candles").Range("H4").ClearContents
'    Set X = Sheets("Candles").Range("I" & Rows.Count).End(xlUp)
'    Sheets("Candles").Range("H4") = X / X.Offset(-Sheets("Candles").Range("L6").Value) - 1
'
'    Sheets("Candles").Range("D2").ClearContents
'    Sheets("Candles").Range("D2") = Sheets("Candles").Range("F3")
'
'    Sheets("Candles").Range("D3").ClearContents
'    Sheets("Candles").Range("D3") = Sheets("Candles").Range("G3")
'
'ElseIf Sheets("Candles").Range("B4").Value <> "DIA" Then
'
'    Sheets("Candles").Range("H4").ClearContents 'last close dow return.
'    Set X = Sheets("Candles").Range("I" & Rows.Count).End(xlUp)
'    Sheets("Candles").Range("H4") = X / X.Offset(-Sheets("Candles").Range("L6").Value) - 1
'
'End If



'With ActiveSheet
    'LastRow = .Cells(.Rows.Count, "I").End(xlUp).Row
    '.Range("H4").Value = .Cells(LastRow, "I").Value / .Cells(LastRow - Range("L6").Value, "I").Value
'End With

'Exit Sub 'Greg Lovern

BadSymbol: 'Greg Lovern
bSymbolNotFound = True
'MsgBox "Symbol " & Symbol & " not found.", vbCritical + vbOKOnly, "Symbol Not Found" 'Greg Lovern
Application.Calculation = xlCalculationAutomatic 'Greg Lovern
Application.DisplayAlerts = True 'Greg Lovern

    For i = 1 To ActiveSheet.QueryTables.Count
        ActiveSheet.QueryTables.Item(1).Delete
    Next i
    
End Sub

Sub CleanUPMakeSureOfTimeAlignmentNYSE()
'NYSE
' CleanUPMakeSureOfTimeAlignmentNYSE Macro
'run AFTER: MakeFirstSetOfLagsNYSE
'does not do UDN
'

'
    Columns("D:G").Select
    Selection.Delete Shift:=xlToLeft
    Range("D1:F1").Select
    Range("D2:F2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("H2:J2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("K:K").Select
    Selection.Delete Shift:=xlToLeft
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "LnCount/LnVolL1"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "VolNormL1"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "HighMinusLowL1"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "VolumeL1"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "NYSEperL1"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "COSpreadL1"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "InterCountDateNormL1"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "InterOpenDateNormL1"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "InterCODateL1"
    Range("K2:M2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "InterTO-YCDateL1"
    Range("N2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "InterNYSEDateL1"
    Range("O2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.SmallScroll ToRight:=86
    ActiveWindow.LargeScroll ToRight:=-6
    Range("A1").Select
End Sub

Sub MakeFieldsNoLagsNYSE()
'
' MakeFieldsNoLagsNYSE Macro

'TO BE USED With NYSE DATA  AFTER SVET (SIMPLE SVET). Next would be to start making lags
'uses Final Row calculation FinalRow = Cells(Rows.Count, 1).End(xlUp).Row
'

'
    FinalRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("C1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.LargeScroll Down:=-12
    Range("C1:C" & FinalRow).Cut Destination:=Range("A1:A" & FinalRow)
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Good"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "COSpread"
    Range("A1").Select
    Selection.End(xlToRight).Select
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]-RC[-6]"
    Range("J2").Select
    Selection.Copy
    
    Range("J2:J" & FinalRow).Select
    ActiveSheet.Paste
    Range("K1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Range("J1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.pasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("K1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "UDN"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]>0,""U"",IF(RC[-1]<0,""D"",""N""))"
    Range("K2").Select
    Selection.Copy
    
    Range("K2:K" & FinalRow).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.pasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("L1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "LnCount/LnVol"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=LN(RC[-10])/LN(RC[-4])"
    Range("L2").Select
    Selection.Copy
    
    Range("L2:L" & FinalRow).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.pasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("M1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "VolNorm"
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=RC[-5]/1000000"
    Range("M2").Select
    Selection.Copy
    
    Range("M2:M" & FinalRow).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.pasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("N1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "HighMinusLow"
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "=RC[-9]-RC[-8]"
    Range("N2").Select
    Selection.Copy
   
    Range("N2:N" & FinalRow).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Application.CutCopyMode = False
    Selection.Copy
    Selection.pasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("O1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "TOMinusYC"
    Range("O3").Select
    ActiveCell.FormulaR1C1 = ""
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "TOMinusYCNotLagged"
    Range("O3").Select
    ActiveCell.FormulaR1C1 = "=RC[-11]-R[-1]C[-8]"
    Range("O3").Select
    Selection.Copy
    
    Range("O3:O" & FinalRow).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Application.CutCopyMode = False
    Selection.Copy
    Selection.pasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    Selection.End(xlToRight).Select
    Range("P1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "InterCountDateNorm"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=RC[-14]*(RC[-15]/4000)"
    Range("P2").Select
    Selection.Copy
    
    Range("P2:P" & FinalRow).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.pasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.NumberFormat = "0.00"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "InterOpenDateNorm"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=RC[-13]*(RC[-16]/4000)"
    Range("Q2").Select
    Selection.Copy
    
    Range("Q2:Q" & FinalRow).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.pasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.NumberFormat = "0.00"
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "InterCODate"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "=(RC[-17]/4000)*RC[-8]"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "=(RC[-17]/40)*RC[-8]"
    Range("R2").Select
    Selection.Copy
    
    Range("R2:R" & FinalRow).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.pasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.NumberFormat = "0.00"
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "InterTO-YCDate"
    Range("S2").Select
    ActiveCell.FormulaR1C1 = ""
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "InterTO-YCDateNotLagged"
    Range("S3").Select
    ActiveCell.FormulaR1C1 = "=RC[-4]*RC[-18]"
    Range("S3").Select
    Selection.Copy
    
    Range("S3:S" & FinalRow).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.pasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.NumberFormat = "0.00"
    Range("T1").Select
    ActiveCell.FormulaR1C1 = "InterNYSEDate"
    Range("T2").Select
    ActiveCell.FormulaR1C1 = "=RC[-11]*(RC[-19]/40)"
    Range("T2").Select
    Selection.Copy
    
    Range("T2:T" & FinalRow).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.NumberFormat = "0.00"
    Selection.Copy
    Application.CutCopyMode = False
    Selection.Copy
    Selection.pasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
End Sub


Sub ExtendedPriceCountMatchSvet()

Dim rowCount As Integer
Dim rowCount2 As Integer
Dim temp As Long
Dim icount As Integer
Dim innerIndex As Integer
Dim ioutIndex As Integer
Dim inIndex As Integer
Dim NrOfColumns As Integer
Dim tmpRange As String, LastCol As String

'Total number of columns
NrOfColumns = ActiveSheet.Rows("1:1").End(xlToRight).Column
'Determine last column letter
LastCol = ActiveSheet.Cells(1, NrOfColumns).Address
LastCol = Replace(LastCol, "$", "")
LastCol = Replace(LastCol, "1", "")

rowCount = WorksheetFunction.CountIf(Columns("A:A"), ">0") - 1
rowCount2 = WorksheetFunction.CountIf(Columns("C:C"), ">0") - 1
icount = 1

Columns("A:A").Select
ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("A1"), _
SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.ActiveSheet.Sort
.SetRange Range(Cells(1, 1), Cells(rowCount, 1))
.Header = xlYes
.MatchCase = False
.Orientation = xlTopToBottom
.SortMethod = xlPinYin
.Apply
End With


For i = 2 To rowCount + 2

    ioutIndex = i
    If Cells(i, 3).Value <> "" And Cells(i, 1).Value <> "" Then
    
        For innerIndex = i To rowCount2 + 2
            inIndex = innerIndex
        
            If Cells(innerIndex, 1).Value = "" Or Cells(innerIndex, 3).Value = "" Then
                Exit For
            End If
   
    
            'ds1.date is sooner than ds2.date 1 day.
            If DateValue(Cells(i, 3).Value) - DateValue(Cells(i, 1).Value) = 1 Then
                If icount > 1 Then
                    Cells(i, 2).Value = (Cells(i, 2).Value + temp) / icount
                End If
                temp = 0
                icount = 1
                Exit For
            End If
            
            'ds1.date is sooner than ds2.date more than 1 day.
            If DateValue(Cells(i, 3).Value) - DateValue(Cells(i, 1).Value) > 1 Then
                temp = temp + Cells(i, 2).Value
                icount = icount + 1
                Range(CStr("A") + CStr(i) + CStr(":B") + CStr(i)).Select
                Selection.Delete Shift:=xlUp
                ioutIndex = ioutIndex - 1
                rowCount = rowCount - 1
                Exit For
            End If
            
            'ds1.date is past than ds2.date.
            If DateValue(Cells(i, 3).Value) - DateValue(Cells(i, 1).Value) < 1 Then
                tmpRange = "C" & CStr(innerIndex) & ":" & LastCol & CStr(innerIndex)
                Range(tmpRange).Select
                Selection.Delete Shift:=xlUp
                innerIndex = innerIndex - 1
                rowCount2 = rowCount2 - 1
            End If
        Next innerIndex
    End If
    i = ioutIndex
Next i

rowCount = WorksheetFunction.CountIf(Columns("A:A"), ">0") - 1
rowCount2 = WorksheetFunction.CountIf(Columns("C:C"), ">0") - 1

For innerIndex2 = rowCount + 3 To rowCount2 + 3

If Cells(innerIndex2, 3).Value <> "" Then
    tmpRange = "C" & CStr(innerIndex2) & ":" & LastCol & CStr(innerIndex2)
    Range(tmpRange).Select
    Selection.Delete Shift:=xlUp
    innerIndex2 = innerIndex2 - 1
End If

Next innerIndex2

    Range("B1").Select
    ActiveCell.FormulaR1C1 = "normCountPrevDaytillMorningNext"
    Application.Run "remove_zero_col_D"
    Range("A1").Select
    Application.Run "cleanLastSpaceFromColumnC"
    Range("A1").Select
    Application.Run " fillInDateInColumnA_AfterSvetWithCurrentDateNYC"
    
    Range("A1").Select
End Sub

Sub MakeFirstSetOfLagsNYSE()
'
' MakeFirstSetOfLags NYSE  Macro
'To be run after MAKEFIELDSNOLAGSNYSE
'

'
    Range("I1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "NYSEper"
    Range("A1").Select
    Selection.End(xlToRight).Select
    Range("U1").Select
    ActiveCell.FormulaR1C1 = "CL1"
    Range("V1").Select
    ActiveCell.FormulaR1C1 = "CL2"
    Range("U1:V1").Select
    Selection.AutoFill Destination:=Range("U1:AN1"), Type:=xlFillDefault
    Range("U1:AN1").Select
    Selection.End(xlToRight).Select
    Range("AO1").Select
    ActiveCell.FormulaR1C1 = "OpenL1"
    Range("AP1").Select
    ActiveCell.FormulaR1C1 = "OpenL2"
    Range("AO1:AP1").Select
    Selection.AutoFill Destination:=Range("AO1:AX1"), Type:=xlFillDefault
    Range("AO1:AX1").Select
    Selection.End(xlToRight).Select
    Range("AY1").Select
    ActiveCell.FormulaR1C1 = "CloseL1"
    Range("AZ1").Select
    ActiveCell.FormulaR1C1 = "CloseL2"
    Range("AY1:AZ1").Select
    Selection.AutoFill Destination:=Range("AY1:BH1"), Type:=xlFillDefault
    Range("AY1:BH1").Select
    Selection.End(xlToRight).Select
    Range("BI1").Select
    ActiveCell.FormulaR1C1 = "HighL1"
    Range("BJ1").Select
    ActiveCell.FormulaR1C1 = "HighL2"
    Range("BI1:BJ1").Select
    Selection.AutoFill Destination:=Range("BI1:BR1"), Type:=xlFillDefault
    Range("BI1:BR1").Select
    Selection.End(xlToRight).Select
    Range("BS1").Select
    ActiveCell.FormulaR1C1 = "LowL1"
    Range("BT1").Select
    ActiveCell.FormulaR1C1 = "LowL2"
    Range("BS1:BT1").Select
    Selection.AutoFill Destination:=Range("BS1:CB1"), Type:=xlFillDefault
    Range("BS1:CB1").Select
    Selection.End(xlToRight).Select
    Range("CC1").Select
    ActiveCell.FormulaR1C1 = "COspreadL1"
    Range("CD1").Select
    ActiveCell.FormulaR1C1 = "COspreadL2"
    Range("CC1:CD1").Select
    Selection.AutoFill Destination:=Range("CC1:CL1"), Type:=xlFillDefault
    Range("CC1:CL1").Select
    Selection.End(xlToRight).Select
    ActiveCell.FormulaR1C1 = "TO-YCL1"
    Range("CM1").Select
    ActiveCell.FormulaR1C1 = "TO-YCL2"
    Range("CL1:CM1").Select
    Selection.AutoFill Destination:=Range("CL1:CU1"), Type:=xlFillDefault
    Range("CL1:CU1").Select
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    ActiveWindow.ScrollRow = 339
    ActiveWindow.ScrollRow = 337
    ActiveWindow.ScrollRow = 336
    ActiveWindow.ScrollRow = 335
    ActiveWindow.ScrollRow = 334
    ActiveWindow.ScrollRow = 333
    ActiveWindow.ScrollRow = 332
    ActiveWindow.ScrollRow = 327
    ActiveWindow.ScrollRow = 326
    ActiveWindow.ScrollRow = 323
    ActiveWindow.ScrollRow = 320
    ActiveWindow.ScrollRow = 318
    ActiveWindow.ScrollRow = 317
    ActiveWindow.ScrollRow = 315
    ActiveWindow.ScrollRow = 313
    ActiveWindow.ScrollRow = 309
    ActiveWindow.ScrollRow = 306
    ActiveWindow.ScrollRow = 303
    ActiveWindow.ScrollRow = 298
    ActiveWindow.ScrollRow = 296
    ActiveWindow.ScrollRow = 295
    ActiveWindow.ScrollRow = 292
    ActiveWindow.ScrollRow = 291
    ActiveWindow.ScrollRow = 290
    ActiveWindow.ScrollRow = 289
    ActiveWindow.ScrollRow = 286
    ActiveWindow.ScrollRow = 283
    ActiveWindow.ScrollRow = 281
    ActiveWindow.ScrollRow = 280
    ActiveWindow.ScrollRow = 278
    ActiveWindow.ScrollRow = 276
    ActiveWindow.ScrollRow = 271
    ActiveWindow.ScrollRow = 267
    ActiveWindow.ScrollRow = 265
    ActiveWindow.ScrollRow = 260
    ActiveWindow.ScrollRow = 255
    ActiveWindow.ScrollRow = 250
    ActiveWindow.ScrollRow = 246
    ActiveWindow.ScrollRow = 241
    ActiveWindow.ScrollRow = 233
    ActiveWindow.ScrollRow = 228
    ActiveWindow.ScrollRow = 226
    ActiveWindow.ScrollRow = 221
    ActiveWindow.ScrollRow = 220
    ActiveWindow.ScrollRow = 217
    ActiveWindow.ScrollRow = 212
    ActiveWindow.ScrollRow = 204
    ActiveWindow.ScrollRow = 197
    ActiveWindow.ScrollRow = 193
    ActiveWindow.ScrollRow = 191
    ActiveWindow.ScrollRow = 184
    ActiveWindow.ScrollRow = 179
    ActiveWindow.ScrollRow = 176
    ActiveWindow.ScrollRow = 171
    ActiveWindow.ScrollRow = 165
    ActiveWindow.ScrollRow = 162
    ActiveWindow.ScrollRow = 160
    ActiveWindow.ScrollRow = 157
    ActiveWindow.ScrollRow = 154
    ActiveWindow.ScrollRow = 145
    ActiveWindow.ScrollRow = 141
    ActiveWindow.ScrollRow = 135
    ActiveWindow.ScrollRow = 130
    ActiveWindow.ScrollRow = 128
    ActiveWindow.ScrollRow = 125
    ActiveWindow.ScrollRow = 120
    ActiveWindow.ScrollRow = 113
    ActiveWindow.ScrollRow = 110
    ActiveWindow.ScrollRow = 107
    ActiveWindow.ScrollRow = 104
    ActiveWindow.ScrollRow = 98
    ActiveWindow.ScrollRow = 95
    ActiveWindow.ScrollRow = 93
    ActiveWindow.ScrollRow = 87
    ActiveWindow.ScrollRow = 81
    ActiveWindow.ScrollRow = 78
    ActiveWindow.ScrollRow = 73
    ActiveWindow.ScrollRow = 70
    ActiveWindow.ScrollRow = 66
    ActiveWindow.ScrollRow = 64
    ActiveWindow.ScrollRow = 63
    ActiveWindow.ScrollRow = 61
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 56
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 51
    ActiveWindow.ScrollRow = 47
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 42
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 38
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    Range("U3").Select
    ActiveSheet.Paste
    Range("V4").Select
    ActiveSheet.Paste
    Range("W5").Select
    ActiveSheet.Paste
    Range("X6").Select
    ActiveSheet.Paste
    Range("Y7").Select
    ActiveSheet.Paste
    Range("Z8").Select
    ActiveSheet.Paste
    Range("AA9").Select
    ActiveSheet.Paste
    Range("AB10").Select
    ActiveSheet.Paste
    Range("AC11").Select
    ActiveSheet.Paste
    Range("AD12").Select
    ActiveSheet.Paste
    Range("AE13").Select
    ActiveSheet.Paste
    Range("AF14").Select
    ActiveSheet.Paste
    Range("AG15").Select
    ActiveSheet.Paste
    Range("AH16").Select
    ActiveSheet.Paste
    Range("AI17").Select
    ActiveSheet.Paste
    Range("AJ18").Select
    ActiveSheet.Paste
    Range("AK19").Select
    ActiveSheet.Paste
    Range("AL20").Select
    ActiveSheet.Paste
    Range("AM21").Select
    ActiveSheet.Paste
    Range("AN22").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("AN1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "CL20"
    Columns("AO:AO").Select
    Range("D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    ActiveWindow.ScrollRow = 339
    ActiveWindow.ScrollRow = 338
    ActiveWindow.ScrollRow = 337
    ActiveWindow.ScrollRow = 336
    ActiveWindow.ScrollRow = 335
    ActiveWindow.ScrollRow = 334
    ActiveWindow.ScrollRow = 333
    ActiveWindow.ScrollRow = 332
    ActiveWindow.ScrollRow = 330
    ActiveWindow.ScrollRow = 327
    ActiveWindow.ScrollRow = 323
    ActiveWindow.ScrollRow = 320
    ActiveWindow.ScrollRow = 317
    ActiveWindow.ScrollRow = 312
    ActiveWindow.ScrollRow = 307
    ActiveWindow.ScrollRow = 304
    ActiveWindow.ScrollRow = 301
    ActiveWindow.ScrollRow = 297
    ActiveWindow.ScrollRow = 292
    ActiveWindow.ScrollRow = 287
    ActiveWindow.ScrollRow = 283
    ActiveWindow.ScrollRow = 272
    ActiveWindow.ScrollRow = 258
    ActiveWindow.ScrollRow = 249
    ActiveWindow.ScrollRow = 239
    ActiveWindow.ScrollRow = 231
    ActiveWindow.ScrollRow = 227
    ActiveWindow.ScrollRow = 224
    ActiveWindow.ScrollRow = 216
    ActiveWindow.ScrollRow = 205
    ActiveWindow.ScrollRow = 200
    ActiveWindow.ScrollRow = 193
    ActiveWindow.ScrollRow = 179
    ActiveWindow.ScrollRow = 167
    ActiveWindow.ScrollRow = 162
    ActiveWindow.ScrollRow = 155
    ActiveWindow.ScrollRow = 153
    ActiveWindow.ScrollRow = 150
    ActiveWindow.ScrollRow = 145
    ActiveWindow.ScrollRow = 141
    ActiveWindow.ScrollRow = 138
    ActiveWindow.ScrollRow = 137
    ActiveWindow.ScrollRow = 136
    ActiveWindow.ScrollRow = 132
    ActiveWindow.ScrollRow = 129
    ActiveWindow.ScrollRow = 127
    ActiveWindow.ScrollRow = 121
    ActiveWindow.ScrollRow = 115
    ActiveWindow.ScrollRow = 113
    ActiveWindow.ScrollRow = 112
    ActiveWindow.ScrollRow = 107
    ActiveWindow.ScrollRow = 104
    ActiveWindow.ScrollRow = 101
    ActiveWindow.ScrollRow = 98
    ActiveWindow.ScrollRow = 97
    ActiveWindow.ScrollRow = 94
    ActiveWindow.ScrollRow = 91
    ActiveWindow.ScrollRow = 90
    ActiveWindow.ScrollRow = 89
    ActiveWindow.ScrollRow = 85
    ActiveWindow.ScrollRow = 82
    ActiveWindow.ScrollRow = 78
    ActiveWindow.ScrollRow = 77
    ActiveWindow.ScrollRow = 75
    ActiveWindow.ScrollRow = 72
    ActiveWindow.ScrollRow = 69
    ActiveWindow.ScrollRow = 68
    ActiveWindow.ScrollRow = 67
    ActiveWindow.ScrollRow = 66
    ActiveWindow.ScrollRow = 63
    ActiveWindow.ScrollRow = 62
    ActiveWindow.ScrollRow = 60
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 56
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 52
    ActiveWindow.ScrollRow = 51
    ActiveWindow.ScrollRow = 49
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 47
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 42
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    Range("AO3").Select
    ActiveSheet.Paste
    Range("AP4").Select
    ActiveSheet.Paste
    Range("AQ5").Select
    ActiveSheet.Paste
    Range("AR6").Select
    ActiveSheet.Paste
    Range("AS7").Select
    ActiveSheet.Paste
    Range("AT8").Select
    ActiveSheet.Paste
    Range("AU9").Select
    ActiveSheet.Paste
    Range("AV10").Select
    ActiveSheet.Paste
    Range("AW11").Select
    ActiveSheet.Paste
    Range("AX12").Select
    ActiveSheet.Paste
    Range("G2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AY3").Select
    ActiveSheet.Paste
    Range("AZ4").Select
    ActiveSheet.Paste
    Range("BA5").Select
    ActiveSheet.Paste
    Range("BB6").Select
    ActiveSheet.Paste
    Range("BC7").Select
    ActiveSheet.Paste
    Range("BD8").Select
    ActiveSheet.Paste
    Range("BE9").Select
    ActiveSheet.Paste
    Range("BF10").Select
    ActiveSheet.Paste
    Range("BG11").Select
    ActiveSheet.Paste
    Range("BH12").Select
    ActiveSheet.Paste
    Range("E2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BI3").Select
    ActiveSheet.Paste
    Range("BJ4").Select
    ActiveSheet.Paste
    Range("BK5").Select
    ActiveSheet.Paste
    Range("BL6").Select
    ActiveSheet.Paste
    Range("BM7").Select
    ActiveSheet.Paste
    Range("BN8").Select
    ActiveSheet.Paste
    Range("BO9").Select
    ActiveSheet.Paste
    Range("BP10").Select
    ActiveSheet.Paste
    Range("BQ11").Select
    ActiveSheet.Paste
    Range("BR12").Select
    ActiveSheet.Paste
    Range("F2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BS3").Select
    ActiveSheet.Paste
    Range("BT4").Select
    ActiveSheet.Paste
    Range("BU5").Select
    ActiveSheet.Paste
    Range("BV6").Select
    ActiveSheet.Paste
    Range("BW7").Select
    ActiveSheet.Paste
    Range("BX8").Select
    ActiveSheet.Paste
    Range("BY9").Select
    ActiveSheet.Paste
    Range("BZ10").Select
    ActiveSheet.Paste
    Range("CA11").Select
    ActiveSheet.Paste
    Range("CB12").Select
    ActiveSheet.Paste
    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CC3").Select
    ActiveSheet.Paste
    Range("CD4").Select
    ActiveSheet.Paste
    Range("CE5").Select
    ActiveSheet.Paste
    Range("CF6").Select
    ActiveSheet.Paste
    Range("CG7").Select
    ActiveSheet.Paste
    Range("CH8").Select
    ActiveSheet.Paste
    Range("CI9").Select
    ActiveSheet.Paste
    Range("CJ10").Select
    ActiveSheet.Paste
    Range("CK11").Select
    ActiveSheet.Paste
    Range("CL12").Select
    Range("O3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CL4").Select
    ActiveSheet.Paste
    Range("CM5").Select
    ActiveSheet.Paste
    Range("CN6").Select
    ActiveSheet.Paste
    Range("CO7").Select
    ActiveSheet.Paste
    Range("CP8").Select
    ActiveSheet.Paste
    Range("CQ9").Select
    ActiveSheet.Paste
    Range("CR10").Select
    ActiveSheet.Paste
    Range("CS11").Select
    ActiveSheet.Paste
    Range("CT12").Select
    ActiveSheet.Paste
    Range("CU13").Select
    ActiveSheet.Paste
    Range("A1").Select
End Sub
 Sub CSV_Import_Loop2()

    Dim sh As Worksheet, sPath As String, sName As String
    Dim ws As Worksheet
    Dim cumFile As String
    Dim CSVs()  As String, i As Long
    
    Set sh = ActiveWorkbook.Sheets("perStockTweets") 'set to current worksheet name
    sPath = "C:\Users\Bill\Desktop\testLongTermTest\"
    sName = Dir(sPath & "*.csv")
    
    '// Store each CSV file in an array
    Do While sName <> ""
        cumFile = sPath & sName
        i = i + 1
        ReDim Preserve CSVs(1 To i)
        CSVs(i) = cumFile
        sName = Dir()
    Loop
    
    If i Then
        'loop the CSVs array
        For i = 1 To UBound(CSVs)
            cumFile = CSVs(i)
            Workbooks.Add Template:= _
                  "C:\Users\Bill\AppData\Roaming\Microsoft\Templates\sizesWithFormV5.xltm" 'for each file in the directory, I want to open it in a new instance of the template.
            Set sh = ActiveWorkbook.Sheets("perStockTweets")
            
            
            With sh.QueryTables.Add(Connection:="TEXT;" & cumFile, Destination:=sh.Range("A1"))
               .TextFileParseType = xlDelimited
               .TextFileCommaDelimiter = True
               .Refresh
            End With
            
            
            'The rest comes for the "DO ALL" button on the form (except for the last line which says sName=Dir()
            'Dim ws As Worksheet
            Set ws = Worksheets("GoogleData")
            MsgBox ("need to set ws.Cells(4,2) etc some way in tweetsLoop")
            ws.Cells(4, 2) = "Coke"
            ws.Cells(3, 2) = Format(Now, "yyyy/mm/dd")
            ws.Cells(2, 2) = "2009-07-15"
            'If TextBox3.Value = "" Then TextBox3.Text = Format(Now, "yyyy/mm/dd")
            
            Application.Run "generateDatesAndMbFormbPerFileFromPerl"
            Application.Run "Macro9"
            Application.Run "fromSvetToCleanUP2"
        Next
    End If

End Sub
 
 
 
 Sub CSV_Import_TweetsLoop()


Dim sh As Worksheet, sPath As String, sName As String
Dim ws As Worksheet
Dim cumFile As String


Set sh = ActiveWorkbook.Sheets("perStockTweets") 'set to current worksheet name
sPath = "C:\Users\Bill\Desktop\testLongTermTest\"
sName = Dir(sPath & "*.csv")
Do While sName <> ""
 'Set cumFile = sPath & sName
 cumFile = sPath & sName
 
  Workbooks.Add Template:= _
        "C:\Users\Bill\AppData\Roaming\Microsoft\Templates\sizesWithFormV5.xltm" 'for each file in the directory, I want to open it in a new instance of the template.
 Set sh = ActiveWorkbook.Sheets("perStockTweets")


  With sh.QueryTables.Add(Connection:="TEXT;" & cumFile, Destination:=sh.Range("A1"))
     .TextFileParseType = xlDelimited
     .TextFileCommaDelimiter = True
     .Refresh
 End With


'The rest comes for the "DO ALL" button on the form (except for the last line which says sName=Dir()
'Dim ws As Worksheet
Set ws = Worksheets("GoogleData")
MsgBox ("need to set ws.Cells(4,2) etc some way in tweetsLoop")
ws.Cells(4, 2) = "Coke"
ws.Cells(3, 2) = Format(Now, "yyyy/mm/dd")
ws.Cells(2, 2) = "2009-07-15"
'If TextBox3.Value = "" Then TextBox3.Text = Format(Now, "yyyy/mm/dd")

Application.Run "generateDatesAndMbFormbPerFileFromPerl"
Application.Run "Macro9"
Application.Run "fromSvetToCleanUP2"

sName = Dir()
 
Loop

 End Sub


Sub AddTimeandChangeToDateOnlyForJapan()
'
' AddTimeandChangeToDateOnly Macro
'

'
FinalRow = Cells(Rows.Count, 2).End(xlUp).Row
    Range("H2").Select
    RngH = "H2:H" & FinalRow
    'note that I normalized by dividing count (for example Coke count)
    'by total number of tweets. But I obtain total by shifting the first 3 of the 6 files
    'to the previous date. But this really means I am counting up to noon of day N as day N-1
    'when counting the total number of tweets (the denominator in the normalization process)
    'so to fix this, I should change this number -10/24 to -12/24. Be careful about
    'daylights savings time.
    
    ActiveCell.FormulaR1C1 = "=RC[-6]- 13/24"
    Range("H2").Select
    Selection.Copy
    Range(RngH).Select
    
    'Range("H2:H61122").Select
    ActiveSheet.Paste
    
     Application.CutCopyMode = False
     Selection.Copy
       Selection.pasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=DATE(YEAR(RC[-1]),MONTH(RC[-1]),DAY(RC[-1]))"
    
   
    'Range(I2).Select
    Selection.Copy
    
      
    
    RngI = "I2:I" & FinalRow
    
    Range(RngI).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.pasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range(RngI).Cut Destination:=Range(RngH)
    Range(RngH).Select
    RngB = "B2:B" & FinalRow
    
    Range(RngH).Cut Destination:=Range(RngB)
    Range("A1").Select
End Sub

Sub kTest()
    
    Const FormulaCol = 7
    
    On Error Resume Next
    Cells(2, FormulaCol).Resize(Cells(Rows.Count, FormulaCol).End(3).Row).SpecialCells(xlCellTypeFormulas, 16).EntireRow.Delete
    
End Sub


Sub Macro9()
      '
      ' Macro9 Macro
      '

      '
10        Range("A1").Select
20        ActiveWorkbook.Worksheets("perStockTweets").Sort.SortFields.Clear
30        ActiveWorkbook.Worksheets("perStockTweets").Sort.SortFields.Add Key:=Range( _
              "B2:B65000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
              xlSortNormal
40        With ActiveWorkbook.Worksheets("perStockTweets").Sort
50            .SetRange Range("A1:G65000")
60            .Header = xlYes
70            .MatchCase = False
80            .Orientation = xlTopToBottom
90            .SortMethod = xlPinYin
100           .Apply
110       End With
120       Application.Run "AddTimeandChangeToDateOnlyForJapan"
          
130       ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
              "perStockTweets!" & Sheets("perStockTweets").Range("A1").CurrentRegion.Address(ReferenceStyle:=xlR1C1), Version:=xlPivotTableVersion14). _
              CreatePivotTable TableDestination:="", TableName:="PivotTable6" _
              , DefaultVersion:=xlPivotTableVersion14

140       With ActiveSheet.PivotTables("PivotTable6").PivotFields("Date")
150           .Orientation = xlRowField
160           .Position = 1
170       End With
180       ActiveSheet.PivotTables("PivotTable6").AddDataField ActiveSheet.PivotTables( _
              "PivotTable6").PivotFields("Tweet ID"), "Sum of Tweet ID", xlSum
      '    Range("B3").Select
190       ActiveCell.Offset(0, 1).Range("A1").Select
200       ActiveSheet.PivotTables("PivotTable6").PivotFields("Sum of Tweet ID").Function _
              = xlCount

'210       Range("A4").Select
'220       Range(Selection, Selection.End(xlToRight)).Select
'230       Range(Selection, Selection.End(xlDown)).Select
'240       Range("A4:B47").Select
Range(Range("A2"), Range("A2").End(xlDown).End(xlToRight).Offset(-1)).Select
250       Selection.Copy
260       Sheets("searchAndMBData").Select
270       Range("A2").Select
280       ActiveCell.Select
290       ActiveSheet.Paste
300       ActiveCell.Offset(-1, 0).Range("A1").Select
310       Sheets("mbPerFileFromPerl").Select
320       Range("A1").Select
330       Application.CutCopyMode = False
340       Range("A2").Select
350       Range(Selection, Selection.End(xlToRight)).Select
360       Range(Selection, Selection.End(xlDown)).Select
370       ActiveWorkbook.Worksheets("mbPerFileFromPerl").Sort.SortFields.Clear
380       ActiveWorkbook.Worksheets("mbPerFileFromPerl").Sort.SortFields.Add Key:=Range _
              ("A2:A65000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
              xlSortNormal
390       With ActiveWorkbook.Worksheets("mbPerFileFromPerl").Sort
400           .SetRange Range("A1:B65000")
410           .Header = xlYes
420           .MatchCase = False
430           .Orientation = xlTopToBottom
440           .SortMethod = xlPinYin
450           .Apply
460       End With
470       Application.Run "deleteNonTweetRecordsFromMbPerFileFromPerl"
480       Range("A1").Select
490       Application.Run "fromMbToNumTweets"
500
510       Application.Run "pushBackTimeForAllTweetsNYTime"
520       Range("D1").Select
530       Range(Selection, Selection.End(xlDown)).Select
540       Selection.NumberFormat = "m/d/yyyy"
550       Range("G1").Select
560       Range(Selection, Selection.End(xlDown)).Select
570       Range(Selection, Selection.End(xlDown)).Select
580       Selection.ClearContents
590       Range("D1").Select
600       ActiveCell.FormulaR1C1 = "dateAdj"
610       Range("A1").Select
620
630       ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
              "mbPerFileFromPerl!" & Sheets("mbPerFileFromPerl").Range("A1").CurrentRegion.Address(ReferenceStyle:=xlR1C1), Version:=xlPivotTableVersion14). _
              CreatePivotTable TableDestination:="", TableName:="PivotTable7" _
              , DefaultVersion:=xlPivotTableVersion14

660       With ActiveSheet.PivotTables("PivotTable7").PivotFields("dateAdj")
670           .Orientation = xlRowField
680           .Position = 1
690       End With
700       ActiveSheet.PivotTables("PivotTable7").AddDataField ActiveSheet.PivotTables( _
              "PivotTable7").PivotFields("countAll"), "Sum of countAll", xlSum
710       Range("A2:B2").Select
720       Range(Selection, Selection.End(xlDown)).Select

Range(Range("A2"), Range("A2").End(xlDown).End(xlToRight).Offset(-1)).Select
740       Selection.Copy
750       Sheets("searchAndMBData").Select
760       Range("D2").Select
770       ActiveSheet.Paste
780       Range("A1").Select
790       Application.CutCopyMode = False
800       Application.Run "normalizeCountsWithCountAll"
810       Range(Selection, Selection.End(xlDown)).Select
820       Selection.Copy
830       Sheets("sheetForSvet").Select
840       Range("A1").Select
850       ActiveSheet.Paste
860       Range("A1").Select
870       Sheets("searchAndMBData").Select
Application.Run "pasteSpecial"
Application.Run "kTest"
880       Range("G1").Select
890       Range(Selection, Selection.End(xlDown)).Select
900       Application.CutCopyMode = False
910       Selection.Copy
920       Sheets("sheetForSvet").Select
930       Range("B1").Select
940       Selection.pasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
              :=False, Transpose:=False
950       Columns("A:A").EntireColumn.AutoFit
960       Range("A1").Select
970       Sheets("GoogleData").Select
980       Application.CutCopyMode = False
990       Application.Run "GetData"
1000      Range("C7").Select
1010      Range(Selection, Selection.End(xlToRight)).Select
1020      Range(Selection, Selection.End(xlDown)).Select
1030      ActiveWorkbook.Worksheets("GoogleData").Sort.SortFields.Clear
1040      ActiveWorkbook.Worksheets("GoogleData").Sort.SortFields.Add Key:=Range( _
              "C8:C65000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
              xlSortNormal
1050      With ActiveWorkbook.Worksheets("GoogleData").Sort
1060          .SetRange Range("C7:I65000")
1070          .Header = xlYes
1080          .MatchCase = False
1090          .Orientation = xlTopToBottom
1100          .SortMethod = xlPinYin
1110          .Apply
1120      End With
            Application.Wait Now() + TimeValue("00:00:02")
1130      Selection.Copy
            Application.Wait Now() + TimeValue("00:00:02")
1140      Sheets("sheetForSvet").Select
1150      Range("C1").Select
1160      Selection.pasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
              :=False, Transpose:=False
1170      Range("C1").Select
1180      Application.CutCopyMode = False
1190      ActiveCell.FormulaR1C1 = "DateForPrice"
1200      Range("C1").Select
1210      Range(Selection, Selection.End(xlDown)).Select
1220      Selection.NumberFormat = "m/d/yyyy"
1230      Range("A1").Select
End Sub

Sub pasteSpecial()
'
' pasteSpecial Macro
'

'
    Range("G1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.pasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub makeRelativePivot2()
'
' makeRelativePivot2 Macro
'

'
   ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "perStockTweets!" & Sheets("perStockTweets").Range("A1").CurrentRegion.Address(ReferenceStyle:=xlR1C1), Version:=xlPivotTableVersion14). _
        CreatePivotTable TableDestination:="", TableName:="PivotTable5" _
        , DefaultVersion:=xlPivotTableVersion14


  
'    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
'        "perStockTweets!R1C1:R3926C7", Version:=xlPivotTableVersion14). _
'        CreatePivotTable TableDestination:="", TableName:="PivotTable5" _
'        , DefaultVersion:=xlPivotTableVersion14

    With ActiveSheet.PivotTables("PivotTable5").PivotFields("Date")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable5").AddDataField ActiveSheet.PivotTables( _
        "PivotTable5").PivotFields("Tweet ID"), "Sum of Tweet ID", xlSum
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveSheet.PivotTables("PivotTable5").PivotFields("Sum of Tweet ID").Function _
        = xlCount
End Sub

Sub CSV_Import()
Dim ws As Worksheet, strFile As String
 
Set ws = ActiveWorkbook.Sheets("Sheet1") 'set to current worksheet name
 
strFile = Application.GetOpenFilename("Text Files (*.csv),*.csv", , "Please selec text file...")
 With ws.QueryTables.Add(Connection:="TEXT;" & strFile, Destination:=ws.Range("A2"))
     .TextFileParseType = xlDelimited
     .TextFileCommaDelimiter = True
     .Refresh
 End With
   ws.Name = "testing"


With ws
.Cells(1, 1) = "Column1"
.Cells(1, 2) = "Column2"
.Cells(1, 3) = "Column3"
.Cells(1, 4) = "Column4"
.Cells(1, 5) = "Column5"
End With

 End Sub


Sub FolderList()
'
' Example Macro to list the files contained in a folder.
'

   Dim x As String, MyName As String
   Dim i As Integer
   Dim Response As Integer, TotalFiles As Integer

   On Error Resume Next

Folder:

   ' Prompt the user for the folder to list.
   x = InputBox(Prompt:="What folder do you want to list?" & vbCr & vbCr _
         & "For example: C:\My Documents", _
         Default:=Options.DefaultFilePath(wdDocumentsPath))

   If x = "" Or x = " " Then
      If MsgBox("Either you did not type a folder name correctly" _
            & vbCr & "or you clicked Cancel. Do you want to quit?" _
            & vbCr & vbCr & _
            "If you want to type a folder name, click No." & vbCr & _
            "If you want to quit, click Yes.", vbYesNo) = vbYes Then
         Exit Sub
      Else
         GoTo Folder
      End If
   End If

   ' Test if folder exists.
   If Dir(x, vbDirectory) = "" Then
      MsgBox "The folder does not exist. Please try again."
      GoTo Folder
   End If

   ' Search the specified folder for files
   ' and type the listing in the document.
   With Application.FileSearch
      .NewSearch
      .FileType = msoFileTypeOfficeFiles
      ' Change the .FileType to the type of files you are looking for;
      ' for example, the following line finds all files:
      ' .FileType = msoFileTypeAllFiles
      .LookIn = x
      .Execute
      TotalFiles = .FoundFiles.Count
      If TotalFiles = 0 Then
         MsgBox ("There are no files in the folder!" & _
               "Please type another folder to list.")
         GoTo Folder
      End If

      ' Create a new document for the file listing.
      Application.Documents.Add
      ActiveDocument.ActiveWindow.View = wdPrintView

      ' Set tabs.
      With Selection.ParagraphFormat.TabStops
         .Add _
               Position:=InchesToPoints(3), _
               Alignment:=wdAlignTabLeft, _
               Leader:=wdTabLeaderSpaces
         .Add _
               Position:=InchesToPoints(4), _
               Alignment:=wdAlignTabLeft, _
               Leader:=wdTabLeaderSpaces
      End With

      ' Type the file list headings.
      Selection.TypeText "File Listing of the "

      With Selection.Font
         .Allcaps = True
         .Bold = True
      End With
      Selection.TypeText x
      With Selection.Font
         .Allcaps = False
         .Bold = False
      End With
      Selection.TypeText " folder!" & vbLf
      Selection.Font.Underline = wdUnderlineSingle
      Selection.TypeText vbLf & "File Name" & vbTab & "File Size" _
            & vbTab & "File Date/Time" & vbLf & vbLf
      Selection.Font.Underline = wdUnderlineNone
      For i = 1 To TotalFiles
         MyName = .FoundFiles(i)
         Selection.TypeText MyName & vbTab & FileLen(MyName) _
               & vbTab & FileDateTime(MyName) & vbLf
      Next i

      ' Type the total number of files found.
      Selection.TypeText vbLf & "Total files in folder = " & TotalFiles & _
            " files."
   End With

   If MsgBox("Do you want to print this folder list?", vbYesNo) = vbYes Then
      Application.ActiveDocument.PrintOut
   End If

   If MsgBox("Do you want to list another folder?", vbYesNo) = vbYes Then
      GoTo Folder
   End If

End Sub

Sub DirLoop()

      Dim MyFile As String, Sep As String

      ' Sets up the variable "MyFile" to be each file in the directory
      ' This example looks for all the files that have an .xls extension.
      ' This can be changed to whatever extension is needed. Also, this
      ' macro searches the current directory. This can be changed to any
      ' directory.

      ' Test for Windows or Macintosh platform. Make the directory request.
      Sep = Application.PathSeparator

      If Sep = "\" Then
         ' Windows platform search syntax.
         MyFile = Dir(CurDir() & Sep & "*.xls")

      Else

         ' Macintosh platform search syntax.
         MyFile = Dir("", MacID("XLS5"))
      End If

      ' Starts the loop, which will continue until there are no more files
      ' found.

      Do While MyFile <> ""

         ' Displays a message box with the name of the file. This can be
         ' changed to any procedure that would be needed to run on every
         ' file in the directory such as opening each file.

         MsgBox CurDir() & Sep & MyFile
         MyFile = Dir()
      Loop

   End Sub

Sub GetFilesInFolder(SourceFolderName As String)

'--- For Example:Folder Name= "D:\Folder Name\"

Dim FSO As Scripting.FileSystemObject
Dim SourceFolder As Scripting.Folder, SubFolder As Scripting.Folder
Dim FileItem As Scripting.File
SourceFolder = "C:\Users\Bill\Desktop\ToolsforTweets\CleanUp\CleanUp\bin\Debug"

    Set FSO = New Scripting.FileSystemObject
    Set SourceFolder = FSO.getfolder(SourceFolderName)

    '--- This is for displaying, whereever you want can be configured

    r = 14
    For Each FileItem In SourceFolder.Files
        Cells(r, 2).Formula = r - 13
        Cells(r, 3).Formula = FileItem.Name
        Cells(r, 4).Formula = FileItem.Path
        Cells(r, 5).Formula = FileItem.Size
        Cells(r, 6).Formula = FileItem.Type
        Cells(r, 7).Formula = FileItem.DateLastModified
        Cells(r, 8).Formula = "=HYPERLINK(""" & FileItem.Path & """,""" & "Click Here to Open" & """)"

        r = r + 1   ' next row number
    Next FileItem

    Set FileItem = Nothing
    Set SourceFolder = Nothing
    Set FSO = Nothing
End Sub
Sub fromSvetToCleanUP2()
Application.Run "ExtendedPriceCountMatchSvet"
Application.Run "MakeFieldsNoLagsNYSE"
Application.Run "MakeFirstSetOfLagsNYSE"
Application.Run "CleanUPMakeSureOfTimeAlignmentNYSE"
Application.Run "CleanUp2ForTrainingSet"





End Sub
