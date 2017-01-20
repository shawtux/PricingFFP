Imports System.IO
Imports System.Resources
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Xml
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel


Public Class Utilities
    Private maxPageFinder As String = "^\s+PAGE\s+([0-9]+)/\s+([0-9]+)"
    Dim parseoLineaTabla As String = "^([0-9]{2}(?#1st matches line number))\s([A-Z0-9]+(?#2nd match farebasis code))\s+([0-9.]+|\s{4,}(?#3rd matches amount rt))+\s{4,}([0-9.]+|\s{4,}(?#4th matches amount ow))\s([A-Z](?#5th matches class))\s+(-|\+|NRF(?#6th matches pen column))\s{1,2}([A-Z]|(?#7th matches empty column. saved just in case))([0-9]{2}[A-Z]{3}|\s+-(?#8th matches dates))\s+([A-Z][0-9]{2}[A-Z]{3}|[0-9]+|-(?#9th matches days))\s*(\+|)\s*([0-9]+|-(?#10th matches AP))(\+|(?#11th matches plus sign))\s*(-|[0-9]+|(?#12 matches min stay))(\+\s*|\s*)([0-9]{1,2}[M]|[0-9]{2,3}|-\s(?#13 matches max column))([A-Z]|\s)([A-Z])"
    Dim paginador As String = ""

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="input">texto de entrada de un FQN y determinar la ultima pagina en caso de haber MD excesivos</param>
    ''' <returns></returns>
    Public Function getCleanFQN(input As String)
        Dim options As RegexOptions = RegexOptions.Singleline Or RegexOptions.Multiline

        Dim m As Match = Regex.Match(input, maxPageFinder, options)
        paginador = m.Groups(2).ToString
        Dim stringArray As String() = input.Split(Environment.NewLine)
        Dim inputLimpio As String = ""
        For Each texto As String In stringArray
            m = Regex.Match(texto, "^\s+PAGE\s+" & paginador & "/\s+" & paginador)
            If m.Success Then
                Exit For
            End If
            m = Regex.Match(texto, "^\s+$")

            If texto.IndexOf("PAGE") < 0 And Not m.Success Then
                inputLimpio = inputLimpio & texto
            End If
        Next

        Return inputLimpio

    End Function

    'Public Function parseFareTable(input As String) As List(Of faredata)
    '    Dim options As RegexOptions = RegexOptions.Singleline Or RegexOptions.Multiline
    '    Dim m As Match = Regex.Match(input, maxPageFinder, options)
    '    paginador = m.Groups(2).ToString
    '    Dim stringArray As String() = input.Split(Environment.NewLine)
    '    Dim arrayFares As New List(Of faredata)
    '    Dim fqnTableRegex As String = "FQD.*?(?=PAGE\s+" & paginador & "/\s+" & paginador & ")"
    '    Console.WriteLine(paginador)
    '    m = Regex.Match(input, fqnTableRegex, options)
    '    'Console.WriteLine(input)
    '    If m.Success Then
    '        'Console.WriteLine(m.Groups(0).Value)
    '        'LN FARE BASIS    OW   USD  RT   B PEN  DATES/DAYS   AP MIN MAXFR
    '        '01 TAWDLA55     10000     20000 T  -     -     -   + -  -  365CR
    '        '08 AOM91TAW     25000     50000 A  -  S  -     -   +21  -  365CR
    '        '09 MEEFX409                 670 M  +     -     -     -   +  6M R
    '        '10 LEELE300                 636 L NRF    -     -     -   +  6M R
    '        '11 NLXSP5GE                 769 N  +  S  -     1234+50+ 7+ 12M M
    '        Dim mc As MatchCollection = Regex.Matches(m.Groups(0).Value, parseoLineaTabla, options)
    '        'Console.WriteLine(mc.Count)
    '        For Each linea As Match In mc
    '            Dim farebasis As faredata = New faredata(linea.Groups(1).Value, linea.Groups(2).Value, linea.Groups(3).Value, linea.Groups(4).Value, linea.Groups(5).Value, linea.Groups(6).Value, linea.Groups(7).Value, linea.Groups(8).Value, linea.Groups(9).Value, linea.Groups(10).Value, linea.Groups(11).Value, linea.Groups(12).Value, linea.Groups(13).Value, linea.Groups(14).Value, linea.Groups(15).Value, linea.Groups(16).Value, linea.Groups(17).Value)
    '            'Console.WriteLine(linea.Groups(0).Value)
    '            arrayFares.Add(farebasis)
    '        Next


    '    Else
    '        Console.WriteLine(m.Success)
    '    End If


    '    Return arrayFares


    '    Console.WriteLine(arrayFares.Count)

    '    Dim xlsWorkBook As Microsoft.Office.Interop.Excel.Workbook
    '    Dim xlsWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
    '    Dim xls As New Microsoft.Office.Interop.Excel.Application

    '    Dim filename As String = "C:\Users\Marcos Orrego\Documents\test.xlsx"



    '    xlsWorkBook = xls.Workbooks.Open(filename)
    '    xlsWorkSheet = xlsWorkBook.Sheets(1)



    '    Dim rownum As Integer = 2
    '    For Each fare In arrayFares
    '        If rownum = 2 Then
    '            xlsWorkSheet.Range(xlsWorkSheet.Cells(rownum - 1, 1), xlsWorkSheet.Cells(rownum - 1, 18)).Value = fare.writeExcelRowHeader
    '            xlsWorkSheet.Range(xlsWorkSheet.Cells(rownum, 1), xlsWorkSheet.Cells(rownum, 18)).Value = fare.writeExcelRow
    '        Else

    '            xlsWorkSheet.Range(xlsWorkSheet.Cells(rownum, 1), xlsWorkSheet.Cells(rownum, 18)).Value = fare.writeExcelRow
    '        End If
    '        rownum = rownum + 1
    '    Next

    '    xlsWorkBook.Save()
    '    xlsWorkBook.Close()
    '    xlsWorkBook = Nothing

    'End Function

    Public Sub parseFQN(input As String, outputFile As String)
        Dim farearray As New List(Of faredata)
        Dim xmldoc As New XmlDocument
        Dim fqn As String = "FQN([0-9]{1,2})\*([0-9]{1,2}).*?(?=FQN|TRANSACTION CODE NOT SUPPORTED)"
        Dim FTC As String = ".*FTC:\s*([A-Z0-9]+[\s\t]*)-.*"

        Dim options As RegexOptions = RegexOptions.Singleline Or RegexOptions.Multiline
        Dim rownum As String = ""
        Dim farebasis As String = ""
        Dim od As String = ""
        xmldoc.LoadXml(My.Resources.ResourceManager.GetObject("Categorias"))

        Dim pasada As Integer = 0
        For Each match As Match In Regex.Matches(input, fqn, options)
            Dim finder As Match
            pasada = pasada + 1
            Dim fareBasisFindandRow As String = "^([0-9]{2})\s([0-9A-Z]+)\s+[0-9]+.*"
            Dim fareBasisOD As String = "([A-Z]{6})/"
            Dim m11 As Match = Regex.Match(match.Value, fareBasisOD, options)

            If m11.Success Then
                od = m11.Groups(1).Value
            End If

            Dim m2 As Match = Regex.Match(match.Value, fareBasisFindandRow, options)
            If m2.Success Then
                rownum = m2.Groups(1).Value
                farebasis = m2.Groups(2).Value

                'Console.WriteLine("rownum: " & rownum)
                'Console.WriteLine("farebasis: " & farebasis)
            End If

            'Console.WriteLine("pasada: " & pasada)
            Dim farebasisRecord As faredata = Nothing

            Dim query = farearray.Where(Function(f As faredata) f.line = rownum And f.farebasis = farebasis And f.OD = od)
            If query.Count() > 0 Then
                farebasisRecord = query.Single
                farearray.Remove(query.Single)
            Else
                Console.WriteLine("query count: " & query.Count())
                Dim matchTabla As Match = Regex.Match(match.Value, parseoLineaTabla, options)
                If matchTabla.Success Then
                    Dim matchOD As Match = Regex.Match(match.Value, " ([A-Z]{6})/", options)
                    farebasisRecord = New faredata(matchTabla.Groups(1).Value, matchTabla.Groups(2).Value, matchTabla.Groups(3).Value, matchTabla.Groups(4).Value, matchTabla.Groups(5).Value, matchTabla.Groups(6).Value, matchTabla.Groups(7).Value, matchTabla.Groups(8).Value, matchTabla.Groups(9).Value, matchTabla.Groups(10).Value, matchTabla.Groups(11).Value, matchTabla.Groups(12).Value, matchTabla.Groups(13).Value, matchTabla.Groups(14).Value, matchTabla.Groups(15).Value, matchTabla.Groups(16).Value, matchTabla.Groups(17).Value, matchOD.Groups(1).Value)

                    Dim accountCode As Match = Regex.Match(match.Value, "/([A-Z]+)$", options)
                    farebasisRecord.accountCode = accountCode.Groups(1).Value
                End If
            End If


            Dim inputLimpio As String = ""
            Dim Categoria As XmlNode = CType(xmldoc.SelectSingleNode("//Categoria[@id=" & match.Groups(2).Value & "]/Texto"), XmlNode)
            Dim strCategoria As String = Categoria.InnerText


            Dim m As Match = Regex.Match(match.Value, maxPageFinder, options)
            If m.Success Then
                paginador = m.Groups(2).ToString
                fqn = strCategoria & ".*?(?=PAGE\s+" & paginador & "/\s+" & paginador & ")"
            Else
                If fqn.Contains("NO MORE PAGE AVAILABLE") Then
                    fqn = strCategoria & ".*?(?=NO MORE PAGE AVAILABLE)"

                End If
            End If
            m = Regex.Match(match.Value, fqn, options)
            If m.Success Then
                Dim stringArray As String() = m.Value.Split(New [Char]() {Environment.NewLine, vbLf})
                For Each texto As String In stringArray
                    If texto.IndexOf("PAGE") < 0 Then
                        inputLimpio = inputLimpio & texto & vbLf
                    End If
                Next


                'TVL RESTRICTION AND SALES RESTRICTION

                If match.Groups(2).Value = 14 Then
                    Dim TVLonBefore As String = "O([0-9]{2}[A-Z]{3}[0-9]{2}|[0-9]{2}[A-Z]{3})"
                    Dim TVLonAfter As String = "E([0-9]{2}[A-Z]{3}[0-9]{2}|[0-9]{2}[A-Z]{3})"
                    Dim completedBy As String = "C([0-9]{2}[A-Z]{3}[0-9]{2}|[0-9]{2}[A-Z]{3})"


                    With farebasisRecord
                        finder = Regex.Match(match.Value, TVLonBefore, options)

                        If finder.Success Then

                            finder = Regex.Match(match.Value, finder.Groups(1).Value & "[0-9]{2}", options)
                            .restrictionBefore = finder.Groups(0).Value
                            ' Console.WriteLine("[" & finder.Groups(0).Value & "]")
                        Else
                            'Console.WriteLine("[" & finder.Groups(0).Value & "]")
                            finder = Regex.Match(match.Value, "COMMENCING ON/BEFORE ([0-9]{2}[A-Z]{3}[0-9]{2}|[0-9]{2}[A-Z]{3}[0-9])")
                            .restrictionAfter = finder.Groups(1).Value

                        End If
                        finder = Regex.Match(match.Value, TVLonAfter, options)
                        If finder.Success Then
                            Console.WriteLine("[" & finder.Groups(0).Value & "]")
                            finder = Regex.Match(match.Value, finder.Groups(1).Value & "[0-9]{2}", options)
                            .restrictionAfter = finder.Groups(0).Value

                        Else
                            finder = Regex.Match(match.Value, "COMMENCING ON/AFTER ([0-9]{2}[A-Z]{3}[0-9]{2}|[0-9]{2}[A-Z]{3}[0-9])")
                            .restrictionAfter = finder.Groups(1).Value
                            Console.WriteLine("[" & finder.Groups(0).Value & "]")
                        End If


                        finder = Regex.Match(match.Value, completedBy, options)
                        If finder.Success Then
                            Dim temp As String = finder.Groups(1).Value
                            finder = Regex.Match(match.Value, finder.Groups(1).Value & "[0-9]{2}", options)
                            If finder.Success Then
                                .completedBy = finder.Groups(0).Value

                            Else
                                .completedBy = temp
                            End If
                            Console.WriteLine("[" & finder.Groups(0).Value & "]")
                        End If

                    End With

                End If


                If match.Groups(2).Value = 15 Then

                    Dim issueonBefore As String = "B([0-9]{2}[A-Z]{3}[0-9]{2}|[0-9]{2}[A-Z]{3})" 'cat15
                    Dim issueonAfter As String = "A([0-9]{2}[A-Z]{3}[0-9]{2}|[0-9]{2}[A-Z]{3})" 'cat15




                    With farebasisRecord

                        finder = Regex.Match(match.Value, issueonAfter, options)
                        If finder.Success Then
                            Dim temp As String = finder.Groups(1).Value
                            finder = Regex.Match(match.Value, finder.Groups(1).Value & "[0-9]{2}", options)
                            If finder.Success Then
                                If .issueAfter = "" Or .issueAfter = Nothing Then
                                    .issueAfter = finder.Groups(0).Value
                                End If
                            Else
                                If .issueAfter = "" Or .issueAfter = Nothing Then
                                    .issueAfter = temp
                                End If

                            End If
                            Console.WriteLine("[" & finder.Groups(0).Value & "]")
                        Else
                            If match.Groups(2).Value = 15 And match.Groups(0).Value.Contains("NO RESTRICTIONS") Then
                                If .issueAfter = "" Or .issueAfter = Nothing Then
                                    .issueAfter = "NO RESTRICTIONS"
                                End If

                            Else
                                finder = Regex.Match(match.Value, "(ISSUED ON/AFTER ([0-9]{2}[A-Z]{3}[0-9]{2}|[0-9]{2}[A-Z]{3}) AND ON/BEFORE\s*([0-9]{2}[A-Z]{3}[0-9]{2})|ISSUED ON/AFTER ([0-9]{2}[A-Z]{3}[0-9]{2}|[0-9]{2}[A-Z]{3}))")
                                For i As Integer = 0 To finder.Groups().Count
                                    Console.WriteLine("issueafter: " & i & " " & finder.Groups(i).Value)
                                Next

                                If (.issueAfter = "" Or .issueAfter = Nothing) Then
                                    .issueAfter = finder.Groups(4).Value
                                End If
                            End If
                        End If

                        finder = Regex.Match(match.Value, issueonBefore, options)
                        If finder.Success Then
                            Dim temp As String = finder.Groups(1).Value
                            finder = Regex.Match(match.Value, finder.Groups(1).Value & "[0-9]{2}", options)
                            .issueBefore = finder.Groups(0).Value
                            Console.WriteLine("[" & finder.Groups(0).Value & "]")
                        Else
                            If match.Groups(2).Value = 15 And match.Groups(0).Value.Contains("NO RESTRICTIONS") Then
                                .issueBefore = "NO RESTRICTIONS"

                            Else
                                finder = Regex.Match(match.Value, "(ISSUED ON/AFTER ([0-9]{2}[A-Z]{3}[0-9]{2}|[0-9]{2}[A-Z]{3}) AND ON/BEFORE\s*([0-9]{2}[A-Z]{3}[0-9]{2})|ON/BEFORE\s*([0-9]{2}[A-Z]{3}[0-9]{2}))")
                                Console.WriteLine("issuebefore: " & finder.Groups(0).Value)
                                If finder.Value.Contains("ISSUED ON/AFTER") Then
                                    .issueBefore = finder.Groups(2).Value
                                Else
                                    .issueBefore = finder.Groups(1).Value
                                End If
                            End If
                        End If

                    End With

                End If










            End If

            'BLACKOUTS
            If match.Groups(2).Value = 11 Then

                Dim noBlackout As String = "NONE"
                finder = Regex.Match(match.Value, noBlackout, options)
                If finder.Success Then
                    farebasisRecord.blackoutOutbound = "NONE"
                    farebasisRecord.blackoutInbound = "NONE"
                Else

                    Dim outbound As String = "FOR OUTBOUND TRAVEL.*?(?=FOR INBOUND TRAVEL)"
                    Dim inbound As String = "FOR INBOUND TRAVEL.*"
                    Console.WriteLine("[" & match.Value & "]")

                    finder = Regex.Match(match.Value, outbound, options)
                    If finder.Success Then
                        Dim outboundDates As String = "([0-9]{2}[A-Z]{3}) THROUGH ([0-9]{2}[A-Z]{3})"
                        For Each dates As Match In Regex.Matches(finder.Value, outboundDates, options)
                            Console.WriteLine(dates.Groups(1).Value & "-" & dates.Groups(2).Value)
                            If farebasisRecord.blackoutOutbound = "" Then
                                farebasisRecord.blackoutOutbound = dates.Groups(1).Value & " - " & dates.Groups(2).Value
                            Else
                                farebasisRecord.blackoutOutbound = farebasisRecord.blackoutOutbound & Chr(10) & dates.Groups(1).Value & " - " & dates.Groups(2).Value
                            End If

                        Next
                    Else
                        Console.WriteLine("outbound Not found")
                    End If
                    finder = Regex.Match(match.Value, inbound, options)
                    If finder.Success Then
                        Dim inboundDates As String = "([0-9]{2}[A-Z]{3}) THROUGH ([0-9]{2}[A-Z]{3})"
                        For Each dates As Match In Regex.Matches(finder.Value, inboundDates, options)
                            If farebasisRecord.blackoutInbound = "" Then
                                farebasisRecord.blackoutInbound = dates.Groups(1).Value & " - " & dates.Groups(2).Value
                            Else
                                farebasisRecord.blackoutInbound = farebasisRecord.blackoutInbound & Chr(10) & dates.Groups(1).Value & " - " & dates.Groups(2).Value
                            End If

                        Next
                    End If

                End If
            End If

            'SEASONS
            If match.Groups(2).Value = 3 Then
                Dim fromRegex As String = "FROM [A-Z]+ TO [A-Z]+.*?(?=TO)"
                Dim validDates As String = "PERMITTED (([0-9]{2}[A-Z]{3}) THROUGH ([0-9]{2}[A-Z]{3})|ON ([0-9]{2}[A-Z]{3}))"
                Dim toRegex As String = "TO [A-Z]+ FROM [A-Z]+.*"

                If inputLimpio.Contains("APPLIES ALL YEAR") Then
                    farebasisRecord.seasonsFrom = "APPLIES ALL YEAR"
                    farebasisRecord.seasonsTo = "APPLIES ALL YEAR"
                Else
                    Dim fromToMatch As Integer = 0
                    Dim fromMatch As Match = Regex.Match(inputLimpio, fromRegex, options)
                    Dim toMatch As Match = Regex.Match(inputLimpio, toRegex, options)
                    If fromMatch.Success Then
                        For Each foundDate As Match In Regex.Matches(fromMatch.Value, validDates, options)

                            If farebasisRecord.seasonsFrom = "" Or farebasisRecord.seasonsFrom = Nothing Then
                                If foundDate.Groups(1).Value.Contains("THROUGH") Then
                                    farebasisRecord.seasonsFrom = foundDate.Groups(2).Value & " - " & foundDate.Groups(3).Value
                                Else
                                    farebasisRecord.seasonsFrom = foundDate.Groups(4).Value
                                End If
                            Else
                                If foundDate.Groups(1).Value.Contains("THROUGH") Then
                                    farebasisRecord.seasonsFrom = farebasisRecord.seasonsFrom & Chr(10) & foundDate.Groups(2).Value & " - " & foundDate.Groups(3).Value
                                Else
                                    farebasisRecord.seasonsFrom = farebasisRecord.seasonsFrom & Chr(10) & foundDate.Groups(4).Value
                                End If

                            End If
                        Next
                    Else
                        fromToMatch = fromToMatch + 1
                    End If
                    If toMatch.Success Then
                        For Each foundDate As Match In Regex.Matches(toMatch.Value, validDates, options)
                            If farebasisRecord.seasonsTo = "" Or farebasisRecord.seasonsTo = Nothing Then
                                If foundDate.Groups(1).Value.Contains("THROUGH") Then
                                    farebasisRecord.seasonsTo = foundDate.Groups(2).Value & " - " & foundDate.Groups(3).Value
                                Else
                                    farebasisRecord.seasonsTo = foundDate.Groups(4).Value
                                End If
                            Else
                                If foundDate.Groups(1).Value.Contains("THROUGH") Then
                                    farebasisRecord.seasonsTo = farebasisRecord.seasonsTo & Chr(10) & foundDate.Groups(2).Value & " - " & foundDate.Groups(3).Value
                                Else
                                    farebasisRecord.seasonsTo = farebasisRecord.seasonsTo & Chr(10) & foundDate.Groups(4).Value
                                End If
                            End If
                        Next
                    Else
                        fromToMatch = fromToMatch + 1
                    End If

                    If fromToMatch = 2 Then

                        For Each foundDate As Match In Regex.Matches(inputLimpio, validDates, options)
                            Console.WriteLine(foundDate.Value)
                            If farebasisRecord.seasonsTo = "" Or farebasisRecord.seasonsTo = Nothing Then
                                If foundDate.Groups(1).Value.Contains("THROUGH") Then
                                    farebasisRecord.seasonsTo = foundDate.Groups(2).Value & " - " & foundDate.Groups(3).Value
                                Else
                                    farebasisRecord.seasonsTo = foundDate.Groups(4).Value
                                End If
                            Else
                                If foundDate.Groups(1).Value.Contains("THROUGH") Then
                                    farebasisRecord.seasonsTo = farebasisRecord.seasonsTo & Chr(10) & foundDate.Groups(2).Value & " - " & foundDate.Groups(3).Value
                                Else
                                    farebasisRecord.seasonsTo = farebasisRecord.seasonsTo & Chr(10) & foundDate.Groups(4).Value
                                End If
                            End If
                            If farebasisRecord.seasonsFrom = "" Or farebasisRecord.seasonsFrom = Nothing Then
                                If foundDate.Groups(1).Value.Contains("THROUGH") Then
                                    farebasisRecord.seasonsFrom = foundDate.Groups(2).Value & " - " & foundDate.Groups(3).Value
                                Else
                                    farebasisRecord.seasonsFrom = foundDate.Groups(4).Value
                                End If
                            Else
                                If foundDate.Groups(1).Value.Contains("THROUGH") Then
                                    farebasisRecord.seasonsFrom = farebasisRecord.seasonsFrom & Chr(10) & foundDate.Groups(2).Value & " - " & foundDate.Groups(3).Value
                                Else
                                    farebasisRecord.seasonsFrom = farebasisRecord.seasonsFrom & Chr(10) & foundDate.Groups(4).Value
                                End If

                            End If

                        Next
                    End If


                End If
            End If

            finder = Regex.Match(match.Value, FTC, options)
            If finder.Success Then
                If farebasisRecord.FTC = "" Or farebasisRecord.FTC = Nothing Then
                    farebasisRecord.FTC = finder.Groups(1).Value
                End If
            Else
                Console.WriteLine("FTC  not found")
            End If
            'Console.WriteLine("group read" & match.Groups(2).Value)
            farearray.Add(farebasisRecord)
        Next



        Dim xlsWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlsWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim xls As New Microsoft.Office.Interop.Excel.Application

        Dim filename As String = outputFile

        If Not File.Exists(filename) Then
            xlsWorkBook = xls.Workbooks.Add()
            xlsWorkBook.SaveAs(filename)
            xlsWorkBook.Close()
        End If

        xlsWorkBook = xls.Workbooks.Open(filename)
        xlsWorkSheet = xlsWorkBook.Sheets(1)



        Dim rownum1 As Integer = 2
        Dim numElements As Integer = 23
        For Each fare In farearray
            If rownum1 = 2 Then
                xlsWorkSheet.Range(xlsWorkSheet.Cells(rownum1 - 1, 1), xlsWorkSheet.Cells(rownum1 - 1, numElements)).Value = fare.writeExcelRowHeader
                xlsWorkSheet.Range(xlsWorkSheet.Cells(rownum1, 1), xlsWorkSheet.Cells(rownum1, numElements)).Value = fare.writeExcelRow

            Else

                xlsWorkSheet.Range(xlsWorkSheet.Cells(rownum1, 1), xlsWorkSheet.Cells(rownum1, numElements)).Value = fare.writeExcelRow



            End If
            rownum1 = rownum1 + 1
        Next

        xlsWorkSheet.Cells.EntireColumn.WrapText = False
        xlsWorkSheet.Cells.EntireColumn.AutoFit()
        xlsWorkSheet.Cells.EntireRow.AutoFit()

        ' Display the range's rows and columns.
        Dim row_min As Integer = xlsWorkSheet.UsedRange.Row
        Dim row_max As Integer = row_min + xlsWorkSheet.UsedRange.Rows.Count - 1
        Dim col_min As Integer = xlsWorkSheet.UsedRange.Column
        Dim col_max As Integer = col_min + xlsWorkSheet.UsedRange.Columns.Count - 1

        For Each cell As Range In xlsWorkSheet.UsedRange
            If cell.Value <> Nothing Then
                If cell.Value.ToString.Contains(Chr(10)) Then
                    cell.WrapText = True
                    cell.EntireColumn.AutoFit()
                End If
            End If
        Next





        xlsWorkBook.Save()
        xlsWorkBook.Close()
        xls.Quit()

        Marshal.ReleaseComObject(xlsWorkSheet)
        Marshal.ReleaseComObject(xlsWorkBook)
        Marshal.ReleaseComObject(xls)


    End Sub






End Class
