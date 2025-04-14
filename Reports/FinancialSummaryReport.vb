Imports ClosedXML.Excel
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Globalization

Public Class FinancialSummaryReport
    Private dbPath As String
    Private tableName As String
    Private outputPath As String
    Private asofdate As String

    Public Sub New(ByVal dbPath As String, ByVal tableName As String, ByVal outputPath As String, ByVal asofdate As String)
        Me.dbPath = dbPath
        Me.tableName = tableName
        Me.outputPath = outputPath
        Me.asofdate = asofdate


    End Sub
    Sub GenerateFinancialSummaryReport()

        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'FinancialSummaryReport

        Dim wbFin As New XLWorkbook()
        Dim wsFin = wbFin.Worksheets.Add("Financial1")
        Dim wsFin2 = wbFin.Worksheets.Add("Financial2")
        ' Define the list of SRRe Asset Types
        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
        Dim dataTable As New DataTable()
        Using conn As New OleDbConnection(connectionString)
            conn.Open()

            ' Query all required data
            Dim query As String = "SELECT * FROM [" & tableName & "] WHERE [As of Date] = ?"

            Using cmd As New OleDbCommand(query, conn)
                ' Add parameter to avoid SQL injection and formatting issues
                cmd.Parameters.AddWithValue("@asofdate", asofdate)

                Using reader As OleDbDataReader = cmd.ExecuteReader()
                    dataTable.Load(reader)
                End Using
            End Using
            Console.WriteLine("Rows in dataTable: " & dataTable.Rows.Count)

            AddReportHeaders(wsFin, "SRRe Asset Type")
            AddReportHeaders(wsFin2, "SRRe Asset Category")

            Dim SRRe_Asset_Fundingquery As String = "SELECT DISTINCT [SRRe Asset Funding] FROM [" & tableName & "]"
            Dim SRRe_Category_query As String = "SELECT DISTINCT [SRRe Asset Category] FROM [" & tableName & "]"
            Dim SRRe_AssetType_query As String = "SELECT DISTINCT [SRRe Asset Type] FROM [" & tableName & "]"

            ' Create a list of lists to store the results
            Dim resultMap As New Dictionary(Of String, List(Of String))()

            Try

                ' Run each query and store the results in the dictionary
                resultMap("SRRe_Asset_Funding") = ExecuteQuery(conn, SRRe_Asset_Fundingquery)
                resultMap("SRRe_AssetType") = ExecuteQuery(conn, SRRe_AssetType_query)
                resultMap("SRRe_Category") = ExecuteQuery(conn, SRRe_Category_query)

            Catch ex As Exception
                Console.WriteLine("Error: " & ex.Message)
            End Try

            Dim TotalOriginal_Face_ValueSum As Double = 0
            Dim TotalCurrent_Face_Sum As Double = 0
            Dim TotalaccruedInterestSum As Double = 0
            Dim TotalMarketValueSum As Double = 0
            Dim TotalStat_Book_ValueSum As Double = 0
            Dim TotalStat_Carry_ValueSum As Double = 0
            Dim TotalYieldtoMaturity As Double = 0
            Dim TotalEffectiveDuration As Double = 0
            Dim TotalSpreadDuration As Double = 0
            Dim TotalOAS As Double = 0
            Dim TotalDiscountMargin As Double = 0
            Dim TotalWeightedAvgLife As Double = 0
            Dim rowNumber As Integer = 2
            Dim ColumnsnamesSheet1() As String = {"SRRe_Asset_Funding", "SRRe_AssetType"}
            DynamicReportGenrations(wsFin, ColumnsnamesSheet1, resultMap, asofdate, dataTable, rowNumber, TotalOriginal_Face_ValueSum, TotalCurrent_Face_Sum, TotalaccruedInterestSum, TotalMarketValueSum, TotalStat_Book_ValueSum,
                TotalStat_Carry_ValueSum)
            WriteTotalSumsToWorksheet(wsFin, rowNumber, TotalOriginal_Face_ValueSum, TotalCurrent_Face_Sum, TotalaccruedInterestSum, TotalMarketValueSum, TotalStat_Book_ValueSum, TotalStat_Carry_ValueSum)

            rowNumber = 2

            Dim columns As Integer() = {11, 12, 13, 14, 15, 16}
            Dim totals As Double() = {TotalYieldtoMaturity, TotalEffectiveDuration, TotalSpreadDuration, TotalOAS, TotalDiscountMargin, TotalWeightedAvgLife}
            CalculateAndWritePercentageRows(wsFin, ColumnsnamesSheet1, resultMap, columns, rowNumber, TotalMarketValueSum, totals)
            WriteTotalsToSheetWithLabels(wsFin, rowNumber, totals, TotalYieldtoMaturity, TotalEffectiveDuration, TotalSpreadDuration, TotalOAS, TotalDiscountMargin, TotalWeightedAvgLife)
            rowNumber = 2
            TotalOriginal_Face_ValueSum = 0
            TotalCurrent_Face_Sum = 0
            TotalaccruedInterestSum = 0
            TotalMarketValueSum = 0
            TotalStat_Book_ValueSum = 0
            TotalStat_Carry_ValueSum = 0
            TotalYieldtoMaturity = 0
            TotalEffectiveDuration = 0
            TotalSpreadDuration = 0
            TotalOAS = 0
            TotalDiscountMargin = 0
            TotalWeightedAvgLife = 0
            totals = New Double() {0, 0, 0, 0, 0, 0}
            Dim ColumnsnamesSheet2() As String = {"SRRe_Asset_Funding", "SRRe_Category"}

            DynamicReportGenrations(wsFin2, ColumnsnamesSheet2, resultMap, asofdate, dataTable, rowNumber, TotalOriginal_Face_ValueSum, TotalCurrent_Face_Sum, TotalaccruedInterestSum, TotalMarketValueSum, TotalStat_Book_ValueSum,
                TotalStat_Carry_ValueSum)
            WriteTotalSumsToWorksheet(wsFin2, rowNumber, TotalOriginal_Face_ValueSum, TotalCurrent_Face_Sum, TotalaccruedInterestSum, TotalMarketValueSum, TotalStat_Book_ValueSum, TotalStat_Carry_ValueSum)

            rowNumber = 2
            CalculateAndWritePercentageRows(wsFin2, ColumnsnamesSheet2, resultMap, columns, rowNumber, TotalMarketValueSum, totals)

            WriteTotalsToSheetWithLabels(wsFin2, rowNumber, totals, TotalYieldtoMaturity, TotalEffectiveDuration, TotalSpreadDuration, TotalOAS, TotalDiscountMargin, TotalWeightedAvgLife)

            ' Extract the directory path
            Dim directoryPath As String = Path.GetDirectoryName(outputPath)
            Dim match As Match = Regex.Match(asofdate, "\d{1,2}/\d{1,2}/\d{4}")
            Dim extractedDate As String = If(match.Success, match.Value, "Date not found")
            Dim newFileName As String = ""
            If extractedDate <> "Date not found" Then
                ' Parse the date using CultureInfo
                Dim parsedDate As DateTime = DateTime.ParseExact(extractedDate, "M/d/yyyy", CultureInfo.InvariantCulture)

                ' Format to MM_dd_yyyy
                Dim formattedDate As String = parsedDate.ToString("MM_dd_yyyy")

                ' Create the new filename
                newFileName = directoryPath & "\Financial Summary" & formattedDate & ".xlsx"

                Console.WriteLine(newFileName)
            Else
                Console.WriteLine("Invalid date format.")
            End If

            Console.WriteLine(newFileName)

            wbFin.SaveAs(newFileName)
            Console.WriteLine("Report generated at: " & outputPath)
        End Using

    End Sub

    Sub WriteTotalsToSheetWithLabels(ws As IXLWorksheet, rowNumber As Integer, totals As Double(), ByRef TotalYieldtoMaturity As Double, ByRef TotalEffectiveDuration As Double, ByRef TotalSpreadDuration As Double, ByRef TotalOAS As Double, ByRef TotalDiscountMargin As Double, ByRef TotalWeightedAvgLife As Double)
        ' Assign totals from array to variables
        TotalYieldtoMaturity = totals(0)
        TotalEffectiveDuration = totals(1)
        TotalSpreadDuration = totals(2)
        TotalOAS = totals(3)
        TotalDiscountMargin = totals(4)
        TotalWeightedAvgLife = totals(5)

        ' Write labels and values to worksheet
        ws.Cell(rowNumber, 11).Value = "TotalYieldtoMaturity"
        ws.Cell(rowNumber + 1, 11).Value = TotalYieldtoMaturity

        ws.Cell(rowNumber, 12).Value = "TotalEffectiveDuration"
        ws.Cell(rowNumber + 1, 12).Value = TotalEffectiveDuration

        ws.Cell(rowNumber, 13).Value = "TotalSpreadDuration"
        ws.Cell(rowNumber + 1, 13).Value = TotalSpreadDuration

        ws.Cell(rowNumber, 14).Value = "TotalOAS"
        ws.Cell(rowNumber + 1, 14).Value = TotalOAS

        ws.Cell(rowNumber, 15).Value = "TotalDiscountMargin"
        ws.Cell(rowNumber + 1, 15).Value = TotalDiscountMargin

        ws.Cell(rowNumber, 16).Value = "TotalWeightedAvgLife"
        ws.Cell(rowNumber + 1, 16).Value = TotalWeightedAvgLife
    End Sub
    Public Sub AddReportHeaders(ws As IXLWorksheet, colName As String)
        ' Add headers to the columns
        ws.Cell(1, 1).Value = "As Of Date"
        ws.Cell(1, 2).Value = "Funding Type"
        ws.Cell(1, 3).Value = colName

        ws.Cell(1, 5).Value = "Sum Original Face Value"
        ws.Cell(1, 6).Value = "Sum Current Face"
        ws.Cell(1, 7).Value = "Sum Accrued Interest"
        ws.Cell(1, 8).Value = "Sum Market Value"
        ws.Cell(1, 9).Value = "Sum Stat Book Value"
        ws.Cell(1, 10).Value = "Sum Stat Carry Value"
        ws.Cell(1, 11).Value = "Yield to Maturity"
        ws.Cell(1, 12).Value = "Effective Duration"
        ws.Cell(1, 13).Value = "Spread Duration"
        ws.Cell(1, 14).Value = "OAS"
        ws.Cell(1, 15).Value = "Discount Margin"
        ws.Cell(1, 16).Value = "Weighted Avg Life"
        ' Optional: Make headers bold
        ws.Range("A1:P1").Style.Font.Bold = True
    End Sub
    Sub WriteTotalSumsToWorksheet(ws As IXLWorksheet, ByRef rowNumber As Integer, TotalOriginal_Face_ValueSum As Decimal, TotalCurrent_Face_Sum As Decimal, TotalaccruedInterestSum As Decimal, TotalMarketValueSum As Decimal, TotalStat_Book_ValueSum As Decimal, TotalStat_Carry_ValueSum As Decimal)
        ws.Cell(rowNumber, 5).Value = "TotalOriginal_Face_ValueSum"
        ws.Cell(rowNumber + 1, 5).Value = TotalOriginal_Face_ValueSum

        ws.Cell(rowNumber, 6).Value = "TotalCurrent_Face_Sum"
        ws.Cell(rowNumber + 1, 6).Value = TotalCurrent_Face_Sum

        ws.Cell(rowNumber, 7).Value = "TotalaccruedInterestSum"
        ws.Cell(rowNumber + 1, 7).Value = TotalaccruedInterestSum

        ws.Cell(rowNumber, 8).Value = "TotalMarketValueSum"
        ws.Cell(rowNumber + 1, 8).Value = TotalMarketValueSum

        ws.Cell(rowNumber, 9).Value = "TotalStat_Book_ValueSum"
        ws.Cell(rowNumber + 1, 9).Value = TotalStat_Book_ValueSum

        ws.Cell(rowNumber, 10).Value = "TotalStat_Carry_ValueSum"
        ws.Cell(rowNumber + 1, 10).Value = TotalStat_Carry_ValueSum

    End Sub
    Sub CalculateAndWritePercentageRows(ws As IXLWorksheet, columName As String(), resultMap As Dictionary(Of String, List(Of String)), columns As Integer(), ByRef rowNumber As Integer, TotalMarketValueSum As Double, ByRef totals As Double())

        For Each fundingType As String In resultMap("SRRe_Asset_Funding")
            For Each loops As String In resultMap(columName(1))
                For i As Integer = 0 To columns.Length - 1
                    Dim col As Integer = columns(i)
                    Dim cellValue As Object = ws.Cell(rowNumber, col).Value
                    Dim numValue As Double = 0

                    If Double.TryParse(cellValue.ToString(), numValue) Then
                        Dim result As Double = 0

                        If TotalMarketValueSum <> 0 Then
                            result = numValue / TotalMarketValueSum

                            ' Avoid NaN or Infinity
                            If Double.IsNaN(result) OrElse Double.IsInfinity(result) Then
                                result = 0
                            End If
                        Else
                            result = 0
                        End If

                        ws.Cell(rowNumber, col).Value = result
                        totals(i) += numValue
                    Else
                        ws.Cell(rowNumber, col).Value = 0
                    End If
                Next

                rowNumber += 1
            Next
        Next
    End Sub
    Public Function DynamicReportGenrations(ws As IXLWorksheet, columName As String(), resultMap As Dictionary(Of String, List(Of String)), asofdate As String, DataTable As DataTable, ByRef rowNumber As Integer, ByRef TotalOriginal_Face_ValueSum As Decimal, ByRef TotalCurrent_Face_Sum As Decimal, ByRef TotalaccruedInterestSum As Decimal, ByRef TotalMarketValueSum As Decimal, ByRef TotalStat_Book_ValueSum As Decimal, ByRef TotalStat_Carry_ValueSum As Decimal)
        Dim Col1 As String = ""
        If columName(1) = "SRRe_AssetType" Then
            Col1 = "SRRe Asset Type"
        Else
            Col1 = "SRRe Asset Category"
        End If

        For Each fundingType As String In resultMap(columName(0)) ' usually only one
            For Each loop2 As String In resultMap(columName(1))

                ' Filter only once
                Dim filteredRows = DataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Asset Funding") = fundingType AndAlso
                                                        r.Field(Of String)(Col1) = loop2 AndAlso
                                                        r.Field(Of String)("As Of Date") = asofdate)

                ' Helper to safely convert
                Dim SafeSum = Function(col As String) As Decimal
                                  Return filteredRows.Sum(Function(r)
                                                              Dim val = r(col).ToString().Trim()
                                                              If String.IsNullOrEmpty(val) OrElse val = "-" OrElse Not IsNumeric(val) Then
                                                                  Return 0D
                                                              End If
                                                              Return Convert.ToDecimal(Convert.ToDouble(val))
                                                          End Function)
                              End Function

                ' Summing values
                Dim Original_Face_ValueSum = SafeSum("Original Face Value")
                Dim Current_Face_Sum = SafeSum("Current Face")
                Dim accruedInterestSum = SafeSum("Accrued Interest")
                Dim MarketValueSum = SafeSum("Market Value")
                Dim Stat_Book_ValueSum = SafeSum("Stat Book Value")
                Dim Stat_Carry_ValueSum = SafeSum("Stat Carry Value")
                Dim YieldtoMaturitysum = SafeSum("Yield to Maturity")
                Dim SpreadDurationsum = SafeSum("Spread Duration")
                Dim EffectiveDurationsum = SafeSum("Effective Duration")
                Dim OASsum = SafeSum("OAS")
                Dim DiscountMarginsum = SafeSum("Discount Margin")
                Dim WeightedAvgLifesum = SafeSum("Weighted Avg Life")

                ' Write to Excel
                ws.Cell(rowNumber, 1).Value = asofdate
                ws.Cell(rowNumber, 2).Value = fundingType
                ws.Cell(rowNumber, 3).Value = loop2
                ws.Cell(rowNumber, 5).Value = Original_Face_ValueSum
                ws.Cell(rowNumber, 6).Value = Current_Face_Sum
                ws.Cell(rowNumber, 7).Value = accruedInterestSum
                ws.Cell(rowNumber, 8).Value = MarketValueSum
                ws.Cell(rowNumber, 9).Value = Stat_Book_ValueSum
                ws.Cell(rowNumber, 10).Value = Stat_Carry_ValueSum
                ws.Cell(rowNumber, 11).Value = YieldtoMaturitysum * MarketValueSum
                ws.Cell(rowNumber, 12).Value = EffectiveDurationsum * MarketValueSum
                ws.Cell(rowNumber, 13).Value = SpreadDurationsum * MarketValueSum
                ws.Cell(rowNumber, 14).Value = OASsum * MarketValueSum
                ws.Cell(rowNumber, 15).Value = DiscountMarginsum * MarketValueSum
                ws.Cell(rowNumber, 16).Value = WeightedAvgLifesum * MarketValueSum

                ' Update totals
                rowNumber += 1
                TotalOriginal_Face_ValueSum += Original_Face_ValueSum
                TotalCurrent_Face_Sum += Current_Face_Sum
                TotalaccruedInterestSum += accruedInterestSum
                TotalMarketValueSum += MarketValueSum
                TotalStat_Book_ValueSum += Stat_Book_ValueSum
                TotalStat_Carry_ValueSum += Stat_Carry_ValueSum

                ' Log for debugging
                Console.WriteLine($"Funding Type: {fundingType}, Asset Type/Asset Category: {loop2}, As Of Date: {asofdate}, Original: {Original_Face_ValueSum}, Current: {Current_Face_Sum}, Accrued: {accruedInterestSum}, Market: {MarketValueSum}, Book: {Stat_Book_ValueSum}, Carry: {Stat_Carry_ValueSum}")
            Next
        Next
    End Function
    Function ExecuteQuery(connection As OleDbConnection, query As String) As List(Of String)
        Dim resultList As New List(Of String)()
        Try
            Using command As New OleDbCommand(query, connection)
                Using reader As OleDbDataReader = command.ExecuteReader()
                    While reader.Read()
                        resultList.Add(reader(0).ToString()) ' Add the first column's value (distinct value)
                    End While
                End Using
            End Using
        Catch ex As OleDbException
            Console.WriteLine("OLE DB Error while executing query: " & ex.Message)
        Catch ex As Exception
            Console.WriteLine("General Error while executing query: " & ex.Message)
        End Try
        Return resultList
    End Function
End Class