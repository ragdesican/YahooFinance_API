

Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq


Public Class Form1
    Dim counter As Integer
    Dim dt As DateTime
    Dim DtSet As System.Data.DataSet
    Dim Crumb As String, Cookie As String, validCookieCrumb As Boolean

    Private Sub Label1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        filltable(dgvMain, "Sheet1$")
        'filltable(dgvQuote, "Quote$")



        'txtDate.Value = Format(Now(), "yyyyMMdd")
        txtVal.Text = "20.00"
        dtp_start.Value = "01/01/2019"



        'For Each row As DataGridViewRow In dgvMain.Rows

        '    row.Cells("Symbol") = New DataGridViewLinkCell
        '    row.Cells("Company") = New DataGridViewLinkCell
        'Next
        TabControl1.TabPages.Remove(TabPage1)

        hidecolumn()

    End Sub

    Sub hidecolumn()
        dgvMain.Columns("Column7").Visible = False
        dgvMain.Columns("Column8").Visible = False
        dgvMain.Columns("Column9").Visible = False
        dgvMain.Columns("AAP").Visible = False
        dgvMain.Columns("Column11").Visible = False
        dgvMain.Columns("Capitalization").Visible = False
        dgvMain.Columns("Level").Visible = False


        Dim g As Integer
        For g = 19 To 28
            dgvMain.Columns(g).Visible = False
        Next

    End Sub
    Sub filltable(dttable As DataGridView, sheetname As String)
        Dim MyConnection As System.Data.OleDb.OleDbConnection

        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
        MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=""Excel 12.0 Xml;HDR=Yes"";Data Source='C:\financeapp\data\raw.xlsx';")
        MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection)
        ' MyCommand.TableMappings.Add("Table", "Net-informations.com")
        DtSet = New System.Data.DataSet

        MyCommand.Fill(DtSet)
        dttable.DataSource = DtSet.Tables(0)
        MyConnection.Close()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Dim nRows As Integer
        Dim val As Object
        nRows = dgvMain.RowCount - 1
        val = InputBox("How many rows download?", "", nRows)
        If Not IsNumeric(val) Then
            Exit Sub
        End If

        nRows = val
        Call getSymbolDetails(nRows)


    End Sub

    Sub getSymbolDetails(nRows As Integer)

        Dim stockCnt As Long, stockArr, stockStr As String, startRw As Long, i As Long, intBlock As Integer
        dt = DateTime.Now

        'dgvQuote.DataSource = Nothing
        'dgvQuote.Rows.Clear()
        Dim xx As Long
        xx = dgvQuote.Rows.Count
        For xx = 1 To dgvQuote.Rows.Count - 1
            dgvQuote.Rows.Remove(dgvQuote.Rows(0))
        Next



        Dim row As Long = 0
        For i = 0 To nRows
            startRw = i
            stockStr = dgvMain.Rows(i).Cells(0).Value
            'dgvQuote.Rows.Add()
            If stockStr Is Nothing Or Trim(stockStr) = "" Then
                Exit For
            End If
            'dgvQuote.DataSource.addnew()

            Call stockPriceAtDate(stockStr, Epoch(GetDate(Trim(txtDate.Value))), i)
            'Call stockPriceAtDate(stockStr, Epoch(GetDate(Trim(txtDate.Text))), i, "1d")
            If dgvMain.Rows(i).Cells(2).Value Is Nothing Then Call stockPriceAtDate(stockStr, Epoch(GetDate(Trim(txtDate.Value))), i, "1d")
            Call stockCloseVolume(stockStr, i)
            Call stock52Week(stockStr, i)
            Call stock50DMA(stockStr, i)



            Dim result As String = stockDownload(stockStr)
            Call splitStock(stockStr, startRw, i, result)

            'AAP
            dgvMain.Rows(i).Cells(9).Value = Format(10000 / dgvMain.Rows(i).Cells(2).Value, "0")
            'Column11'Historical Gain
            If IsError(dgvMain.Rows(i).Cells(3).Value / dgvMain.Rows(i).Cells(2).Value) Then
                dgvMain.Rows(i).Cells(10).Value = 0
                dgvMain.Rows(i).Cells(11).Value = 0
            Else
                dgvMain.Rows(i).Cells(10).Value = Format(100 * (dgvMain.Rows(i).Cells(3).Value / dgvMain.Rows(i).Cells(2).Value - 1), "0")
                dgvMain.Rows(i).Cells(11).Value = Format(100 * (dgvMain.Rows(i).Cells(3).Value / dgvMain.Rows(i).Cells(2).Value - 1), "0.00")
            End If
            'level
            Dim E4 As Double = IIf(IsDBNull(dgvMain.Rows(i).Cells(4).Value), 0, dgvMain.Rows(i).Cells(4).Value)
            Dim D4 As Double = IIf(IsDBNull(dgvMain.Rows(i).Cells(3).Value), 0, dgvMain.Rows(i).Cells(3).Value)
            Dim G4 As Double = IIf(IsDBNull(dgvMain.Rows(i).Cells(6).Value), 0, dgvMain.Rows(i).Cells(6).Value)

            Dim F4 As Double = IIf(IsDBNull(dgvMain.Rows(i).Cells(5).Value), 0, dgvMain.Rows(i).Cells(5).Value)
            Dim F2 As Double = CDbl(Trim(txtVal.Text))


            If E4 > G4 And E4 > D4 Then
                dgvMain.Rows(i).Cells(13).Value = Format(100 * (D4 - G4) / (E4 - G4), "0.00")
            Else
                If D4 >= E4 Then
                    dgvMain.Rows(i).Cells(13).Value = 100
                Else
                    dgvMain.Rows(i).Cells(13).Value = ""
                End If
            End If

            'cyclegain
            If IsError(E4 / F4) Then
                dgvMain.Rows(i).Cells(16).Value = 0
            Else
                dgvMain.Rows(i).Cells(16).Value = Format(100 * (E4 / F4 - 1), "0")
            End If
            Dim Q4 As Double = dgvMain.Rows(i).Cells(16).Value
            'DV


            If Q4 > 50 Then
                If D4 > E4 Then
                    dgvMain.Rows(i).Cells(14).Value = 100 * (D4 / E4 - 1)
                Else
                    If D4 < F4 Then
                        dgvMain.Rows(i).Cells(14).Value = 100 * (D4 / F4 - 1)
                    Else
                        dgvMain.Rows(i).Cells(14).Value = ""
                    End If
                End If
            Else
                dgvMain.Rows(i).Cells(14).Value = ""
            End If

            'levels
            If D4 > F4 Then
                dgvMain.Rows(i).Cells(18).Value = Format(100 * (D4 - F4) / (E4 - F4), "0.00")
            Else
                If D4 > F4 Then
                    dgvMain.Rows(i).Cells(18).Value = Format(100 * (D4 - F4) / (E4 - F4), "0.00")
                Else
                    dgvMain.Rows(i).Cells(18).Value = 0
                End If
            End If
            Dim S4 As Double = dgvMain.Rows(i).Cells(18).Value

            'Signal

            If Q4 > F2 And S4 >= 100 Then
                dgvMain.Rows(i).Cells(15).Value = 1
            Else
                If Q4 > F2 And S4 <= 0 Then
                    dgvMain.Rows(i).Cells(15).Value = -1
                Else
                    dgvMain.Rows(i).Cells(15).Value = ""
                End If
            End If



            'nowgain
            If D4 >= E4 Or D4 = 0 Then
                dgvMain.Rows(i).Cells(17).Value = 0
            Else
                If D4 < E4 Then
                    dgvMain.Rows(i).Cells(17).Value = Format(100 * (E4 / D4 - 1), "0.00")
                End If
            End If

            stockStr = ""

            label_status.Text = "Getting data from Yahoo Finance " & " - Row " & i & " of " & nRows & " in " & CInt(DateTime.Now.Subtract(dt).TotalSeconds.ToString) & " seconds"
            My.Application.DoEvents()
            'dgvQuote.Rows(i).Cells(1).Value = Sheets("Quotes").Range("B" & i)
            'Sheets("Main").Range("D" & i) = Sheets("Quotes").Range("C" & i)
            'Sheets("Main").Range("E" & i).Formula = Val(Sheets("Quotes").Range("D" & i))
            'Sheets("Main").Range("F" & i).Formula = Val(Sheets("Quotes").Range("E" & i))
            'Sheets("Main").Range("G" & i).Formula = Val(Sheets("Quotes").Range("F" & i))
            'Sheets("Main").Range("H" & i).Formula = Val(Sheets("Quotes").Range("N" & i))

        Next
exitSub:

        label_status.Text = ""
        labeltimer.Text = ""
        MsgBox("Loaded in " & CInt(DateTime.Now.Subtract(dt).TotalSeconds.ToString) & " seconds")

    End Sub



    Sub stockPriceAtDate(stk As String, t1epoch As Long, row As Long, Optional freq As String = "1h")
        Dim httpObject As Object
        httpObject = CreateObject("Msxml2.ServerXMLHTTP.6.0")
        Dim jsonText As String
        Dim jsonObject As Object
        Dim sourceArr, splitArr, splitArr2(1, 15), colCnt As Integer, copyArr
        Dim Item As Object
        Dim itm As Object
        Dim i As Integer
        Dim reStr As String
        Dim surl As String

        'freq = "1h"
        Dim t2epoch As Long
        t2epoch = 259200
        t2epoch = t1epoch + t2epoch

        On Error GoTo NoStock
        surl = "https://query1.finance.yahoo.com/v8/finance/chart/" & stk & "?symbol=" & stk & "&period1=" & t1epoch & "&period2=" & t2epoch & "&interval=" & freq & ""
        'sRequest = surl
        httpObject.Open("GET", surl, False)
        httpObject.Send
        jsonText = httpObject.responsetext

        If jsonText = "" Then
NoStock:


        Else
            Dim v As String
            Dim json As JObject = JObject.Parse(jsonText)
            'dgvQuote.Rows(row).Cells(1).Value = Val(json.SelectToken("chart.result[0].indicators.quote[0].close[0]"))
            dgvMain.Rows(row).Cells(2).Value = Format(Val(json.SelectToken("chart.result[0].indicators.quote[0].close[0]")), "0.00")
        End If
    End Sub
    Function GetDate(Instring As String) As Date
        Dim newDate As Date
        Dim tmpDate As String
        Dim tmpDay As String
        Dim tmpMonth As String
        Dim tmpYear As String

        'Sample Date
        '20130105

        'Remove the Day of the Week.
        tmpDay = Trim(Mid(Instring, 4, 2))

        'Get the month"
        tmpMonth = Trim(Mid(Instring, 1, 2))

        'Get the year"
        tmpYear = Trim(Mid(Instring, 7, 4))

        'Convert string to date
        'newDate = DateValue(tmpMonth & "/" & tmpDay & "/" & tmpYear)
        newDate = DateSerial(tmpYear, tmpMonth, tmpDay)

        GetDate = newDate
    End Function
    Function Epoch(d As Date) As Long
        Epoch = DateDiff("s", #1/1/1970#, d)
    End Function

    Sub stockCloseVolume(stk As String, row As Long)
        Dim httpObject As Object
        httpObject = CreateObject("Msxml2.ServerXMLHTTP.6.0")
        Dim jsonText As String
        Dim jsonObject As Object
        'Dim sourceArr, splitArr, splitArr2(1, 15), colCnt As Integer, copyArr
        Dim Item As Object
        Dim copyArr

        Dim surl As String
        On Error GoTo NoStock
        surl = "https://query1.finance.yahoo.com/v8/finance/chart/" & stk & "?symbol=" & stk & "&interval=1d"
        'Debug.Print sURL

        httpObject.Open("GET", surl, False)
        httpObject.Send
        jsonText = httpObject.responsetext

        If jsonText = "" Then
NoStock:
            Dim Res As String = "404"
            'Worksheets("Main").Range("R" & row) = "First/Last Stock Price Not Found."
        Else
            'MsgBox jsonText

            Dim json As JObject = JObject.Parse(jsonText)
            'dgvQuote.Rows(row).Cells(1).Value = Val(json.SelectToken("chart.result[0].indicators.quote[0].close[0]"))

            'dgvQuote.Rows(row).Cells(2).Value = Val(json.SelectToken("chart.result[0].indicators.quote[0].close[0]"))
            dgvMain.Rows(row).Cells(3).Value = Format(Val(json.SelectToken("chart.result[0].indicators.quote[0].close[0]")), "0.00")
            'dgvQuote.Rows(row).Cells(13).Value = Val(json.SelectToken("chart.result[0].indicators.quote[0].volume[0]"))
            dgvMain.Rows(row).Cells(7).Value = Val(json.SelectToken("chart.result[0].indicators.quote[0].volume[0]"))
            'dgvMain.Rows(row).Cells(3).Value
            'For Each Item In json.SelectToken("chart.result[0]")

            '    dgvQuote.Rows(row).Cells(3).Value = Item("indicators")("quote")(1)("close")(1)
            '    dgvQuote.Rows(row).Cells(14).Value = Item("indicators")("quote")(1)("volume")(1)
            '    '       Debug.Print Item("indicators")("quote")(1)("close")(1)
            '    '       Debug.Print Item("indicators")("quote")(1)("volume")(1)

            '    'i = i + 1

            'Next

        End If

    End Sub

    Sub stock52Week(stk As String, row As Long)

        Dim httpObject As Object
        httpObject = CreateObject("Msxml2.ServerXMLHTTP.6.0")
        Dim jsonText As String
        Dim jsonObject As Object
        'Dim sourceArr, splitArr, splitArr2(1, 15), colCnt As Integer, copyArr
        Dim Item As Object
        Dim highArr(), lowArr() As Double
        Dim i As Integer
        Dim j As Integer
        ' Dim freq As String: freq = "1d"
        Dim freq As String : freq = "3mo"

        '    Dim stk As String: stk = "nvr"
        '    Dim row As Long: row = 4

        Dim today As Date : today = Now
        Dim back52Wks As Date : back52Wks = DateAdd("ww", -52, today)

        Dim epochNow As Long : epochNow = DateDiff("s", #1/1/1970#, today)
        Dim epoch52Back As Long : epoch52Back = DateDiff("s", #1/1/1970#, back52Wks)

        On Error GoTo NoStock
        Dim surl As String = "https://query1.finance.yahoo.com/v8/finance/chart/" & stk & "?symbol=" & stk & "&period1=" & epoch52Back & "&period2=" & epochNow & "&interval=" & freq & ""
        'Debug.Print sURL
        'sRequest = surl
        httpObject.Open("GET", surl, False)
        httpObject.Send
        jsonText = httpObject.responsetext

        If jsonText = "" Then
NoStock:
            'res = "404"
            'Worksheets("Main").Range("R" & row) = "First/Last Stock Price Not Found."
        Else
            'MsgBox jsonText

            Dim json As JObject = JObject.Parse(jsonText)
            Dim child As Object




            'For Each child In json.SelectToken("chart.result[0]")
            i = 0
            j = 0

            Do Until i > json.SelectToken("chart.result[0].indicators.quote[0].high").Count - 1
                ReDim Preserve highArr(j)
                ReDim Preserve lowArr(j)
                highArr(j) = Val(json.SelectToken("chart.result[0].indicators.quote[0].high[" & i & "]"))
                lowArr(j) = Val(json.SelectToken("chart.result[0].indicators.quote[0].low[" & i & "]"))
                i = i + 1
                j = j + 1
            Loop
            'Debug.Print Item("indicators")("quote")(1)("close").Count
            '        Debug.Print UBound(highArr)
            '        Debug.Print UBound(lowArr)
            '        Debug.Print "52 Week high : " & Format(WorksheetFunction.Max(highArr), "0.00")
            '        Debug.Print "52 Week Low : " & Format(WorksheetFunction.Min(lowArr), "0.00")
            'dgvQuote.Rows(row).Cells(3).Value = Format(highArr.Max, "0.00")
            'dgvQuote.Rows(row).Cells(4).Value = Format(lowArr.Min, "0.00")

            dgvMain.Rows(row).Cells(4).Value = Format(highArr.Max, "0.00")
            dgvMain.Rows(row).Cells(5).Value = Format(lowArr.Min, "0.00")
            'Next

        End If

    End Sub

    Sub stock50DMA(stk As String, row As Long)

        Dim httpObject As Object
        httpObject = CreateObject("Msxml2.ServerXMLHTTP.6.0")
        Dim jsonText As String
        Dim jsonObject As Object
        'Dim sourceArr, splitArr, splitArr2(1, 15), colCnt As Integer, copyArr
        Dim Item As Object
        Dim closeArr() As Double
        Dim i As Integer
        Dim j As Integer
        Dim freq As String : freq = "1d"

        '    Dim stk As String: stk = "vfc"
        '    Dim row As Long: row = 4

        Dim today As Date : today = Now
        Dim back50Ds As Date : back50Ds = DateAdd("d", -100, today)

        Dim epochNow As Long : epochNow = DateDiff("s", #1/1/1970#, today)
        Dim epochback50Ds As Long : epochback50Ds = DateDiff("s", #1/1/1970#, back50Ds)

        On Error GoTo NoStock
        Dim surl As String = "https://query1.finance.yahoo.com/v8/finance/chart/" & stk & "?symbol=" & stk & "&period1=" & epochback50Ds & "&period2=" & epochNow & "&interval=" & freq & ""
        'Debug.Print sURL
        'sRequest = surl
        httpObject.Open("GET", surl, False)
        httpObject.Send
        jsonText = httpObject.responsetext

        If jsonText = "" Then
NoStock:
            'res = "404"
            'Worksheets("Main").Range("R" & row) = "First/Last Stock Price Not Found."
        Else
            'MsgBox jsonText
            Dim json As JObject = JObject.Parse(jsonText)
            Dim ValsSum As Double
            'For Each Item In jsonObject("chart")("result")

            i = json.SelectToken("chart.result[0].indicators.quote[0].close").Count - 1
            j = 0
            If (i > 50) Then
                Do Until j = 50
                    ValsSum = ValsSum + CDbl(json.SelectToken("chart.result[0].indicators.quote[0].close[" & i - 1 & "]"))
                    'ReDim Preserve closeArr(j)

                    'closeArr(j) = val(Item("indicators")("quote")(1)("close")(i))
                    'Debug.Print Item("indicators")("quote")(1)("close").Count
                    'Debug.Print Item("indicators")("quote")(1)("close")(i)
                    i = i - 1
                    j = j + 1
                Loop
                'dgvQuote.Rows(row).Cells(5).Value = Format(ValsSum / 50, "0.00")

                dgvMain.Rows(row).Cells(6).Value = Format(ValsSum / 50, "0.00")
                ' Worksheets("Quotes").Range("F" & row) = Format(WorksheetFunction.Sum(closeArr) / 50, "0.00")
            Else
                dgvQuote.Rows(row).Cells(15).Value = "Not enough data for 50 Day MA Calculation."
            End If

            'Next

        End If

    End Sub

    Function stockDownload(strSmbl As String) As String
        Dim testurl As String
        Dim url As String
        Dim i As Integer
        Dim httpObject As Object
        httpObject = CreateObject("Msxml2.ServerXMLHTTP.6.0")
        Dim jsonText As String
        Dim jsonObject As Object
        Dim Item As Object


        strSmbl = Replace(strSmbl, "GOLD", "SGOL")
        strSmbl = Replace(strSmbl, "NASDAQ", "^IXIC")
        Dim Prefix As String = "https://query1.finance.yahoo.com/v10/finance/quoteSummary/" + strSmbl + "?modules=financialdata,defaultKeyStatistics,cashflowStatementHistoryQuarterly"
        testurl = Prefix


        On Error GoTo NoStock
        httpObject.Open("GET", testurl, False)
        httpObject.Send
        jsonText = httpObject.responsetext

        If jsonText = "" Then
NoStock:
            'res = "404"
            'Worksheets("Main").Range("R" & row) = "First/Last Stock Price Not Found."
        Else
            stockDownload = jsonText
        End If


        '    With Worksheets("Quotes").QueryTables.Add(Connection:=
        '    "URL;" & testurl _
        '    , Destination:=rDest)
        '        .Name = "Conn" + strSmbl
        '        .FieldNames = True
        '        .RowNumbers = False
        '        .FillAdjacentFormulas = False
        '        .PreserveFormatting = True
        '        .RefreshOnFileOpen = False
        '        .BackgroundQuery = True
        '        .RefreshStyle = xlOverwriteCells ' xlInsertDeleteCells
        '        .SavePassword = False
        '        .SaveData = True
        '        .AdjustColumnWidth = True
        '        .RefreshPeriod = 0
        '        .WebSelectionType = xlEntirePage
        '        .WebFormatting = xlWebFormattingNone
        '        .WebPreFormattedTextToColumns = True
        '        .WebConsecutiveDelimitersAsOne = True
        '        .WebSingleBlockTextImport = False
        '        .WebDisableDateRecognition = False
        '        .WebDisableRedirections = False
        '        .Refresh BackgroundQuery:=False
        'End With
    End Function

    Sub splitStock(stockStr As String, startRw As Long, endRw As Long, json As String)
        'Dim ws As Worksheet
        Dim jsonText As String
        Dim jsonObject As Object
        Dim sourceArr, splitArr, splitArr2(1, 15), colCnt As Integer, copyArr
        Dim Item As Object

        'sourceArr = Worksheets("Quotes").Range("AA" & startRw & ":AB" & endRw)

        'jsonText = Worksheets("Quotes").Range("AA" & startRw & ":AB" & endRw)

        'For i = 1 To endRw - startRw + 1

        jsonText = json
        Dim jsonObj As JObject = JObject.Parse(jsonText)

        On Error Resume Next
        'For Each Item In jsonObject("quoteSummary")("result")

        'Debug.Print stockStr & ":" & Item("financialData")("currentPrice")("raw")

        dgvQuote.Rows(startRw).Cells(0).Value = stockStr
        'Worksheets("Quotes").Range("C" & startRw) = Item("financialData")("currentPrice")("raw")
        dgvMain.Rows(startRw).Cells(19).Value = jsonObj.SelectToken("quoteSummary.result[0].financialData.targetMeanPrice.raw")

        dgvMain.Rows(startRw).Cells(20).Value = jsonObj.SelectToken("quoteSummary.result[0].defaultKeyStatistics.trailingEps.raw")

        dgvMain.Rows(startRw).Cells(22).Value = jsonObj.SelectToken("quoteSummary.result[0].defaultKeyStatistics.forwardPE.raw")
        dgvMain.Rows(startRw).Cells(23).Value = jsonObj.SelectToken("quoteSummary.result[0].defaultKeyStatistics.pegRatio.raw")
        If jsonObj.SelectToken("quoteSummary.result[0].defaultKeyStatistics.priceToBook.raw") Is Nothing Then
            dgvMain.Rows(startRw).Cells(24).Value = 0
        Else
            dgvMain.Rows(startRw).Cells(24).Value = jsonObj.SelectToken("quoteSummary.result[0].defaultKeyStatistics.priceToBook.raw")
        End If

        Dim entvalue As Double = CDbl(jsonObj.SelectToken("quoteSummary.result[0].defaultKeyStatistics.enterpriseValue.raw"))
        Dim totrev As Double = CDbl(jsonObj.SelectToken("quoteSummary.result[0].financialData.totalRevenue.raw"))
        dgvMain.Rows(startRw).Cells(25).Value = Format(entvalue / totrev, "0.00")

        dgvMain.Rows(startRw).Cells(27).Value = CDbl(jsonObj.SelectToken("quoteSummary.result[0].defaultKeyStatistics.sharesOutstanding.raw")) * CDbl(jsonObj.SelectToken("quoteSummary.result[0].financialData.currentPrice.raw"))

        ' If Item("cashflowStatementHistoryQuarterly")("cashflowStatements")(1)("dividendsPaid")("raw") = "" Then
        'Worksheets("Quotes").Range("I" & startRw) = ""
        'Else

        dgvMain.Rows(startRw).Cells(21).Value = Math.Abs(CDbl(jsonObj.SelectToken("quoteSummary.result[0].cashflowStatementHistoryQuarterly.cashflowStatements[0].dividendsPaid.raw")) / CDbl(jsonObj.SelectToken("quoteSummary.result[0].defaultKeyStatistics.sharesOutstanding.raw")))
        'End If

        'Debug.Print stockStr & ":" & Item("defaultKeyStatistics")("sharesOutstanding")("raw")
        'Debug.Print stockStr & ":" & Abs(Item("cashflowStatementHistoryQuarterly")("cashflowStatements")(1)("dividendsPaid")("raw")) / Item("defaultKeyStatistics")("sharesOutstanding")("raw")

        ' i = i + 1

        'Next
        '        splitArr = Split(sourceArr(i, 1), ",")
        '        colCnt = UBound(splitArr)
        '        If colCnt = 14 Theni
        '            Worksheets("Quotes").Range("A" & startRw + i - 1 & ":O" & startRw + i - 1) = splitArr
        '        Else
        '            For j = 0 To 14
        '                splitArr2(0, j) = vbNullString
        '            Next
        '            For j = 0 To colCnt
        '                If j <= 12 Then
        '                    splitArr2(0, 14 - j) = splitArr(colCnt - j)
        '                ElseIf j <> colCnt Then
        '                    splitArr2(0, 1) = splitArr(colCnt - j) + splitArr2(0, 1)
        '                ElseIf j = colCnt Then
        '                    splitArr2(0, 0) = splitArr(0)
        '                End If
        '            Next
        '            Worksheets("Quotes").Range("A" & startRw + i - 1 & ":O" & startRw + i - 1) = splitArr2
        '        End If
        'Next

        'copyArr = Worksheets("Second Input").Range("D" & startRw & ":O" & endRw)
        'For i = 1 To endRw - startRw + 1
        '    For j = 1 To 13
        '        If copyArr(i, j) = "N/A" Or copyArr(i, j) = 0 Then
        '            copyArr(i, j) = vbNullString
        '        End If
        '    Next
        'Next
        ' Worksheets("Main").Range("E" & startRw & ":P" & endRw) = copyArr
    End Sub

    Sub createchart(val As String)
        Dim startDate As Date = dtp_start.Value
        Dim endDate As Date = dtp_end.Value
        On Error GoTo Err
        Dim Attempt As Integer, resultFromYahoo As String
        dt = DateTime.Now
        Call getCookieCrumb()
        If validCookieCrumb = False Then
            MsgBox("Something seems to be wrong! Please check internet connection and try again", vbCritical, "Error")
            Exit Sub
        End If

        Attempt = 1
        Do
            resultFromYahoo = getYahooFinanceData(Trim(val),
                        CStr(DateDiff("s", "1/1/1970", Format(startDate, "MM/dd/yyyy"))),
                        CStr(DateDiff("s", "1/1/1970", Format(endDate, "MM/dd/yyyy"))),
                        "1d", Cookie, Crumb)
            If Mid(resultFromYahoo, 1, 41) = "Date,Open,High,Low,Close,Adj Close,Volume" Then Exit Do
            Debug.Print(resultFromYahoo)
            Attempt = Attempt + 1
        Loop While Attempt <= 3

        If Mid(resultFromYahoo, 1, 41) <> "Date,Open,High,Low,Close,Adj Close,Volume" Then
            MsgBox("Could not return data! Please make sure inputs are valid and try again.", vbExclamation)
            Exit Sub
        End If

        Dim objApp As Object
        Dim oBook As Excel.Workbook
        Dim oSheet As Excel.Worksheet
        objApp = CreateObject("Excel.Application")
        oBook = objApp.WorkBooks.Open("C:\financeapp\data\chart.xlsx")
        objApp.visible = False
        oSheet = oBook.Worksheets(1)


        PopulateData(resultFromYahoo, oSheet)
        updatecharts(oSheet, oBook, objApp)

        txtSymbol.Text = val



        oBook.Close(False)
        oBook = Nothing
        objApp.Quit()
        objApp = Nothing

        label_status.Text = "Ready"
        MsgBox("Loaded in " & CInt(DateTime.Now.Subtract(dt).TotalSeconds.ToString) & " seconds")
        Exit Sub
Err:
        MsgBox("Error! " & Err.Number & Err.Description)

        'oBook.Close(False)
        'oBook = Nothing
        'objApp.Quit()
        'objApp = Nothing

        'Dim s As New DataVisualization.Charting.Series
        's.Name = "aline"
        's.ChartType = DataVisualization.Charting.SeriesChartType.Line



        'For index As Integer = 1 To 10
        '    s.Points.AddXY("1990", 27)
        '    s.Points.AddXY("1991", 15)
        '    s.Points.AddXY("1992", index)
        'Next

        'Chart1.Series.Add(s)

    End Sub

    Private Sub dgvMain_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvMain.CellContentClick
        If e.ColumnIndex = 1 Then
            If MsgBox("Opening Website.. Proceed?", vbYesNo, "Website") = vbYes Then
                Process.Start("http://www.nasdaq.com/symbol/" & Trim(dgvMain.Rows.Item(e.RowIndex).Cells(0).Value))
            End If

            'MsgBox("company")
        ElseIf e.ColumnIndex = 0 Then
            'createchart(Trim(dgvMain.Rows.Item(e.RowIndex).Cells(0).Value))
            If MsgBox("Creating Chart.. Proceed?", vbYesNo, "Chart") = vbYes Then
                label_status.Text = "Creating Chart... Please Wait."
                createchartver2(Trim(dgvMain.Rows.Item(e.RowIndex).Cells(0).Value))
                TabControl1.SelectedIndex = 1
                label_status.Text = "Ready"
            End If

        End If
    End Sub

    Sub getCookieCrumb()
        Dim i As Integer
        Dim str As String
        Dim crumbStartPos As Long
        Dim crumbEndPos As Long
        Dim objRequest

        validCookieCrumb = False

        On Error GoTo ErrCrumb
        For i = 0 To 5
            objRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
            With objRequest
                .Open("GET", "https://finance.yahoo.com/lookup?s=blahblah", False)
                .setRequestHeader("Content-Type", "application/x-www-form-urlencoded; charset=UTF-8")
                '            Application.ScreenUpdating = False
                .Send
                ' Application.ScreenUpdating = True
                .waitForResponse(10)
                Cookie = Split(.getResponseHeader("Set-Cookie"), ";")(0)
                crumbStartPos = InStrRev(.responsetext, """crumb"":""") + 9
                crumbEndPos = crumbStartPos + 11
                Crumb = Mid(.responsetext, crumbStartPos, crumbEndPos - crumbStartPos)
            End With

            If Len(Crumb) = 11 Then
                validCookieCrumb = True
                Exit For
            End If
        Next i

        objRequest = Nothing
        Exit Sub

ErrCrumb:
        objRequest = Nothing

        MsgBox("Error! Probably internet connection lost. " & Err.Number & ". " & Err.Description)
        End
    End Sub

    Private Function getYahooFinanceData(stockTicker As String, startDate As String, endDate As String, frequency As String, Cookie As String, Crumb As String) As String
        Dim objRequest
        Dim tickerURL As String

        tickerURL = "https://query1.finance.yahoo.com/v7/finance/download/" & stockTicker &
                "?period1=" & startDate & "&period2=" & endDate &
                "&interval=" & frequency & "&events=history" & "&crumb=" & Crumb


        objRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
        With objRequest
            .Open("GET", tickerURL, False)
            .setRequestHeader("Cookie", Cookie)
            .Send
            .waitForResponse
            getYahooFinanceData = .responsetext
        End With

        objRequest = Nothing
    End Function

    Private Sub TabPage3_Click(sender As Object, e As EventArgs) Handles TabPage3.Click

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Sub PopulateData(strResult As String, oSheet As Excel.Worksheet)
        Dim csv_rows() As String, iRows As Long, CSV_Fields() As String, rngCls As Excel.Range, LRw As Long

        csv_rows = Split(strResult, Chr(10))

        With oSheet
            .Range("A3:B" & oSheet.Rows.Count).ClearContents()
            rngCls = .Range("B3")
            For iRows = LBound(csv_rows) + 1 To UBound(csv_rows)
                If Trim(csv_rows(iRows)) <> "" Then
                    CSV_Fields = Split(csv_rows(iRows), ",")
                    If IsDate(CSV_Fields(0)) And IsNumeric(CSV_Fields(5)) Then
                        rngCls.Offset(0, -1).Value = CDate(CSV_Fields(0))
                        rngCls.Value = Val(CSV_Fields(5))
                    End If
                End If
                rngCls = rngCls.Offset(1)

                label_status.Text = "Creating Chart " & " - Data " & iRows & " of " & UBound(csv_rows) & " in " & CInt(DateTime.Now.Subtract(dt).TotalSeconds.ToString) & " seconds"
            Next

            LRw = .Cells(oSheet.Rows.Count, 1).End(Excel.XlDirection.xlUp).row
            If LRw > 3 Then
                .Range(.Cells(LRw, 1).Offset(1), .Cells(oSheet.Rows.Count, "E")).ClearContents()
                .Range(.Cells(oSheet.Rows.Count, 3).End(Excel.XlDirection.xlUp), .Cells(LRw, "E")).FillDown()
            End If

            .Sort.SortFields.Clear()
            .Sort.SortFields.Add(Key:=oSheet.Range("A2:A" & oSheet.Rows.Count),
                        SortOn:=Excel.XlSortOn.xlSortOnValues, Order:=Excel.XlSortOrder.xlDescending, DataOption:=Excel.XlSortDataOption.xlSortNormal)

            With .Sort
                .SetRange(oSheet.Range("A2:B" & oSheet.Rows.Count))
                .Header = Excel.XlYesNoGuess.xlYes
                .MatchCase = False
                .Orientation = Excel.Constants.xlTopToBottom
                .SortMethod = Excel.XlSortMethod.xlPinYin
                .Apply()
            End With

        End With




    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Trim(txtSymbol.Text) = "" Then
            Exit Sub
        End If
        createchartver2(Trim(txtSymbol.Text))
        TabControl1.SelectedIndex = 1
    End Sub

    Sub updatecharts(shChart As Excel.Worksheet, oBook As Excel.Workbook, xlapp As Excel.Application) '(shChart As Worksheet)
        Dim lastrow As Long, min As Double, y As Double ', shData As Worksheet

        '    Set shData = Worksheets("data")
        With shChart

            lastrow = .Cells(.Rows.Count, "A").End(Excel.XlDirection.xlUp).row
            .ChartObjects(1).Chart.SetSourceData(Source:= .Range("A3:B" & lastrow))

            min = xlapp.WorksheetFunction.Min(.Range("B3:B" & lastrow))

            txtMin.Text = Format(min, "0.00")
            txtMax.Text = Format(xlapp.WorksheetFunction.Max(.Range("B3:B" & lastrow)), "0.00")
            y = 0 ' .ChartObjects(1).Chart.Axes(xlValue).MinimumScale
            Do While min > y
                y = y + .ChartObjects(1).Chart.Axes(Excel.XlAxisType.xlValue).MajorUnit
            Loop
            '        Debug.Print y
            .ChartObjects(1).Chart.Axes(Excel.XlAxisType.xlValue).MinimumScale = y - .ChartObjects(1).Chart.Axes(Excel.XlAxisType.xlValue).MajorUnit
            '        .ChartObjects(1).Chart.Axes(xlValue).MinimumScale = 20 ' min - .ChartObjects(1).Chart.Axes(xlCategory).MajorUnit Mod min
            'Second chart
            Dim MyDataSource1 As Excel.Range
            Dim MyDataSource2 As Excel.Range
            MyDataSource1 = .Range("A3:A" & lastrow)
            MyDataSource2 = .Range("C3:C" & lastrow)
            .ChartObjects(2).Chart.SetSourceData(Source:=xlapp.Union(MyDataSource1, MyDataSource2))


            'oBook.Activate()
            '.ChartObjects(2).CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap, Excel.XlPictureAppearance.xlScreen)
            Clipboard.Clear()
            .ChartObjects(1).CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap)

            PictureBox1.Image = CType(Clipboard.GetData(DataFormats.Bitmap), Bitmap)
            Clipboard.Clear()

            .ChartObjects(2).CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap)
            PictureBox2.Image = CType(Clipboard.GetData(DataFormats.Bitmap), Bitmap)
            '.ChartObjects(0).export("C:\Users\EDGARNACIS\Desktop\upwork\MyExcelChart.png", "PNG")
            'Dim bmp As New Bitmap(200, 200)
            'Dim rec As New Rectangle(0, 0, 200, 200)
            '.ChartObjects(1).Chart.DrawToBitmap(bmp, rec)
            'PictureBox1.Image = bmp
            ''{
            ''    chart1.SaveImage(MS, ChartImageFormat.Bmp);
            ''    Bitmap bm = New Bitmap(MS);
            ''    Clipboard.SetImage(bm);
            ''}
            '            .ChartObjects(1).chartarea.select
            '                PictureBox1


        End With
    End Sub

    Sub createchartver2(val As String)
        dt = DateTime.Now
        Dim objApp As Object
        Dim oBook As Excel.Workbook
        Dim oSheet As Excel.Worksheet
        objApp = CreateObject("Excel.Application")
        oBook = objApp.Workbooks.Open("C:\financeapp\data\chart.xlsm")
        objApp.Visible = False
        oSheet = oBook.Worksheets(1)
        oSheet.Range("J2").Value = val
        oSheet.Range("N1").Value = Format(dtp_start.Value, "MM/dd/yyyy")
        oSheet.Range("N2").Value = Format(dtp_end.Value, "MM/dd/yyyy")

        objApp.Run("chart")

        Clipboard.Clear()
        oSheet.ChartObjects(1).CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap)

        PictureBox1.Image = CType(Clipboard.GetData(DataFormats.Bitmap), Bitmap)
        Clipboard.Clear()

        oSheet.ChartObjects(2).CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap)
        PictureBox2.Image = CType(Clipboard.GetData(DataFormats.Bitmap), Bitmap)

        txtSymbol.Text = val
        txtMax.Text = Format(oSheet.Range("Q2").Value, "0.00")
        txtMin.Text = Format(oSheet.Range("T2").Value, "0.00")

        oBook.Close(False)
        oBook = Nothing
        objApp.Quit()
        objApp = Nothing


        MsgBox("Loaded in " & CInt(DateTime.Now.Subtract(dt).TotalSeconds.ToString) & " seconds")
    End Sub

    Public Sub PasteData(ByRef dgv As DataGridView)
        Dim tArr() As String
        Dim arT() As String
        Dim i, ii As Integer
        Dim c, cc, r As Integer
        Dim MyNewRow As DataRow
        tArr = Clipboard.GetText().Split(Environment.NewLine)

        r = dgv.RowCount
        c = dgv.SelectedCells(0).ColumnIndex
        For i = 0 To tArr.Length - 2
            arT = tArr(i).Split(vbTab)
            cc = c
            MyNewRow = DtSet.Tables(0).NewRow
            For ii = 0 To arT.Length - 1



                With MyNewRow
                    .Item(ii) = arT(ii).TrimStart
                    '.Item(1) = 1234

                End With


                cc = cc + 1
            Next
            DtSet.Tables(0).Rows.Add(MyNewRow)
            DtSet.Tables(0).AcceptChanges()
            r = r + 1
        Next

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        PasteData(dgvMain)
    End Sub

    Private Sub txtVal_TextChanged(sender As Object, e As EventArgs) Handles txtVal.TextChanged

    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        If MsgBox("Do you want to save?", vbYesNo, "SAVE") = vbYes Then
            label_status.Text = "Saving.... Please Wait"
            Dim columnsCount As Integer = dgvMain.Columns.Count

            Dim objApp As Object
            Dim oBook As Excel.Workbook
            Dim oSheet As Excel.Worksheet
            objApp = CreateObject("Excel.Application")
            oBook = objApp.Workbooks.Open("C:\financeapp\data\raw.xlsx")
            objApp.Visible = False
            oSheet = oBook.Worksheets(1)
            oSheet.Range("A1:AA20000").Clear()

            For Each column In dgvMain.Columns
                oSheet.Cells(1, column.Index + 1).Value = column.Name
            Next


            For i As Integer = 0 To dgvMain.Rows.Count - 2
                Dim columnIndex As Integer = 0
                Do Until columnIndex = columnsCount
                    If dgvMain.Item(columnIndex, i).Value.ToString Is Nothing Then
                        oSheet.Cells(i + 2, columnIndex + 1).Value = ""
                    Else
                        oSheet.Cells(i + 2, columnIndex + 1).Value = dgvMain.Item(columnIndex, i).Value.ToString
                    End If

                    columnIndex += 1
                Loop
            Next

            oBook.Close(True)
            oBook = Nothing
            objApp.Quit()
            objApp = Nothing
        End If



    End Sub

    Private Sub dgvMain_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles dgvMain.CellFormatting

        On Error GoTo err
        If e.ColumnIndex = 2 Or e.ColumnIndex = 3 Or e.ColumnIndex = 4 Or e.ColumnIndex = 5 Or e.ColumnIndex = 6 _
            Or e.ColumnIndex = 7 Or e.ColumnIndex = 8 Or e.ColumnIndex = 9 Or e.ColumnIndex = 10 Or e.ColumnIndex = 11 _
            Or e.ColumnIndex = 12 Or e.ColumnIndex = 13 Then
            If Not dgvMain.Rows.Item(e.RowIndex).Cells(e.ColumnIndex).Value Is Nothing Then
                If dgvMain.Rows.Item(e.RowIndex).Cells(e.ColumnIndex).Value < 0 Then
                    dgvMain.Rows.Item(e.RowIndex).Cells(e.ColumnIndex).Style.ForeColor = Color.Red
                End If
            End If
        End If
        'For i As Integer = 0 To dgvMain.Rows.Count - 1
        '    If dgvMain.Rows(i).Cells(0).Value < 0 Then
        '        dgvMain.Rows(i).Cells(0).Style.ForeColor = Color.Red
        '    End If
        'Next
        Exit Sub
Err:

    End Sub
End Class
