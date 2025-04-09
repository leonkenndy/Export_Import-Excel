Imports System.Data.OleDb
Imports ClosedXML.Excel
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Globalization
Imports DocumentFormat.OpenXml.Spreadsheet
Imports System.ComponentModel
Public Class FinancialSummaryReport
    Private dbPath As String
    Private tableName As String
    Private outputPath As String
    Private asofdate As String
    Private rowIndex As Integer

    Public Sub New(ByVal dbPath As String, ByVal tableName As String, ByVal outputPath As String, ByVal asofdate As String)
        Me.dbPath = dbPath
        Me.tableName = tableName
        Me.outputPath = outputPath
        Me.asofdate = asofdate
        Me.rowIndex = 0

    End Sub
    Sub GenerateFinancialSummaryReport()

        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'FinancialSummaryReport

        Dim wbFin As New XLWorkbook()
        Dim wsFin = wbFin.Worksheets.Add("Report")
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


            ' Add headers to the columns
            wsFin.Cell(1, 1).Value = "Funding Type"
            'wsFin.Cell(1, 2).Value = "Coupon Type"
            wsFin.Cell(1, 2).Value = "Category"
            wsFin.Cell(1, 3).Value = "SRRe Asset Type"
            wsFin.Cell(1, 4).Value = "As Of Date"
            wsFin.Cell(1, 5).Value = "Sum Original Face Value"
            wsFin.Cell(1, 6).Value = "Sum Current Face"
            wsFin.Cell(1, 7).Value = "Sum Accrued Interest"
            wsFin.Cell(1, 8).Value = "Sum Market Value"
            wsFin.Cell(1, 9).Value = "Sum Stat Book Value"
            wsFin.Cell(1, 10).Value = "Sum Stat Carry Value"
            wsFin.Cell(1, 11).Value = "Yield to Maturity"
            wsFin.Cell(1, 12).Value = "Effective Duration"
            wsFin.Cell(1, 13).Value = "Spread Duration"
            wsFin.Cell(1, 14).Value = "OAS"
            wsFin.Cell(1, 15).Value = "Discount Margin"
            wsFin.Cell(1, 16).Value = "Weighted Avg Life"
            ' Queries to select distinct values from different columns
            Dim SRRe_Asset_Fundingquery As String = "SELECT DISTINCT [SRRe Asset Funding] FROM [" & tableName & "]"
            'Dim Coupon_Type_query As String = "SELECT DISTINCT [SRRe Coupon Type] FROM [" & tableName & "]"
            Dim SRRe_Category_query As String = "SELECT DISTINCT [SRRe Asset Category] FROM [" & tableName & "]"
            Dim SRRe_Rating_query As String = "SELECT DISTINCT [SRRe Asset Type] FROM [" & tableName & "]"

            ' Create a list of lists to store the results
            Dim resultMap As New Dictionary(Of String, List(Of String))()

            ' Create a connection to the database


            Try

                ' Run each query and store the results in the dictionary
                resultMap("SRRe_Asset_Funding") = ExecuteQuery(conn, SRRe_Asset_Fundingquery)
                'resultMap("Coupon_Type") = ExecuteQuery(conn, Coupon_Type_query)
                resultMap("SRRe_Rating") = ExecuteQuery(conn, SRRe_Rating_query)
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



            Dim rowNumber As Integer = 2
            ' Example loop through each combination
            For Each fundingType As String In resultMap("SRRe_Asset_Funding")
                For Each assetCategory As String In resultMap("SRRe_Category")
                    For Each ratingNumber As String In resultMap("SRRe_Rating")

                        ' Filter the data with specific conditions
                        Dim Original_Face_ValueSum As Decimal = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Asset Funding") = fundingType AndAlso
                                                                                                 r.Field(Of String)("SRRe Asset Category") = assetCategory AndAlso
                                                                                                 r.Field(Of String)("SRRe Asset Type") = ratingNumber AndAlso
                                                                                                 r.Field(Of String)("As Of Date") = asofdate
                                                                                                 ).Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Original Face Value"))))

                        Dim Current_Face_Sum As Decimal = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Asset Funding") = fundingType AndAlso
                                                                                                 r.Field(Of String)("SRRe Asset Category") = assetCategory AndAlso
                                                                                                 r.Field(Of String)("SRRe Asset Type") = ratingNumber AndAlso
                                                                                                 r.Field(Of String)("As Of Date") = asofdate
                                                                                                 ).Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Current Face"))))
                        Dim accruedInterestSum As Decimal = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Asset Funding") = fundingType AndAlso
                                                                                                 r.Field(Of String)("SRRe Asset Category") = assetCategory AndAlso
                                                                                                 r.Field(Of String)("SRRe Asset Type") = ratingNumber AndAlso
                                                                                                 r.Field(Of String)("As Of Date") = asofdate
                                                                                                 ).Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Accrued Interest"))))
                        Dim MarketValueSum As Decimal = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Asset Funding") = fundingType AndAlso
                                                                                                 r.Field(Of String)("SRRe Asset Category") = assetCategory AndAlso
                                                                                                 r.Field(Of String)("SRRe Asset Type") = ratingNumber AndAlso
                                                                                                 r.Field(Of String)("As Of Date") = asofdate
                                                                                                 ).Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Market Value"))))
                        Dim Stat_Book_ValueSum As Decimal = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Asset Funding") = fundingType AndAlso
                                                                                                 r.Field(Of String)("SRRe Asset Category") = assetCategory AndAlso
                                                                                                 r.Field(Of String)("SRRe Asset Type") = ratingNumber AndAlso
                                                                                                 r.Field(Of String)("As Of Date") = asofdate
                                                                                                 ).Sum(Function(r)
                                                                                                           Dim value As String = r("Stat Book Value").ToString().Trim()
                                                                                                           If String.IsNullOrEmpty(value) Then
                                                                                                               Return 0D ' Treat blanks as zero
                                                                                                           Else
                                                                                                               Return Convert.ToDecimal(Convert.ToDouble(value))
                                                                                                           End If
                                                                                                       End Function)
                        Dim Stat_Carry_ValueSum As Decimal = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Asset Funding") = fundingType AndAlso
                                                                                                 r.Field(Of String)("SRRe Asset Category") = assetCategory AndAlso
                                                                                                 r.Field(Of String)("SRRe Asset Type") = ratingNumber AndAlso
                                                                                                 r.Field(Of String)("As Of Date") = asofdate
                                                                                                 ).Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

                        Dim YieldtoMaturitysum As Decimal = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Asset Funding") = fundingType AndAlso
                                                                                                 r.Field(Of String)("SRRe Asset Category") = assetCategory AndAlso
                                                                                                 r.Field(Of String)("SRRe Asset Type") = ratingNumber AndAlso
                                                                                                 r.Field(Of String)("As Of Date") = asofdate
                                                                                                 ).Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Yield to Maturity"))))
                        Dim SpreadDurationsum As Decimal = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Asset Funding") = fundingType AndAlso
                                                                                                 r.Field(Of String)("SRRe Asset Category") = assetCategory AndAlso
                                                                                                 r.Field(Of String)("SRRe Asset Type") = ratingNumber AndAlso
                                                                                                 r.Field(Of String)("As Of Date") = asofdate
                                                                                                 ).Sum(Function(r)
                                                                                                           Dim value As String = r("Spread Duration").ToString()
                                                                                                           Dim numericValue As Decimal
                                                                                                           If value <> "-" AndAlso Decimal.TryParse(value, numericValue) Then
                                                                                                               Return numericValue
                                                                                                           Else
                                                                                                               Return 0
                                                                                                           End If
                                                                                                       End Function)
                        Dim EffectiveDurationsum As Decimal = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Asset Funding") = fundingType AndAlso
                                                                                                  r.Field(Of String)("SRRe Asset Category") = assetCategory AndAlso
                                                                                                  r.Field(Of String)("SRRe Asset Type") = ratingNumber AndAlso
                                                                                                  r.Field(Of String)("As Of Date") = asofdate).
                                                                                Sum(Function(r)
                                                                                        Dim value As String = r("Effective Duration").ToString().Trim()
                                                                                        Dim numericValue As Decimal

                                                                                        ' Check if value is numeric and not "-" or "'_"'
                                                                                        If value <> "-" AndAlso Decimal.TryParse(value, numericValue) Then
                                                                                            Return numericValue
                                                                                        Else
                                                                                            Return 0
                                                                                        End If
                                                                                    End Function)

                        Dim OASsum As Decimal = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Asset Funding") = fundingType AndAlso
                                                                                                 r.Field(Of String)("SRRe Asset Category") = assetCategory AndAlso
                                                                                                 r.Field(Of String)("SRRe Asset Type") = ratingNumber AndAlso
                                                                                                 r.Field(Of String)("As Of Date") = asofdate
                                                                                                 ).Sum(Function(r)
                                                                                                           Dim value As String = r("OAS").ToString()
                                                                                                           If String.IsNullOrEmpty(value) OrElse Not IsNumeric(value) Then
                                                                                                               Return 0
                                                                                                           Else
                                                                                                               Return Convert.ToDecimal(Convert.ToDouble(value))
                                                                                                           End If
                                                                                                       End Function)
                        Dim DiscountMarginsum As Decimal = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Asset Funding") = fundingType AndAlso
                                                                                                 r.Field(Of String)("SRRe Asset Category") = assetCategory AndAlso
                                                                                                 r.Field(Of String)("SRRe Asset Type") = ratingNumber AndAlso
                                                                                                 r.Field(Of String)("As Of Date") = asofdate
                                                                                                 ).Sum(Function(r)
                                                                                                           Dim value As String = r("Discount Margin").ToString()
                                                                                                           If String.IsNullOrEmpty(value) OrElse Not IsNumeric(value) Then
                                                                                                               Return 0
                                                                                                           Else
                                                                                                               Return Convert.ToDecimal(Convert.ToDouble(value))
                                                                                                           End If
                                                                                                       End Function)
                        Dim WeightedAvgLifesum As Decimal = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Asset Funding") = fundingType AndAlso
                                                                                                 r.Field(Of String)("SRRe Asset Category") = assetCategory AndAlso
                                                                                                 r.Field(Of String)("SRRe Asset Type") = ratingNumber AndAlso
                                                                                                 r.Field(Of String)("As Of Date") = asofdate
                                                                                                 ).Sum(Function(r)
                                                                                                           Dim value As String = r("Weighted Avg Life").ToString()
                                                                                                           If String.IsNullOrEmpty(value) OrElse Not IsNumeric(value) Then
                                                                                                               Return 0
                                                                                                           Else
                                                                                                               Return Convert.ToDecimal(Convert.ToDouble(value))
                                                                                                           End If
                                                                                                       End Function)
                        ' Output the result for the current combination
                        ' Write the data to the next row

                        wsFin.Cell(rowNumber, 1).Value = fundingType

                        wsFin.Cell(rowNumber, 2).Value = assetCategory
                        wsFin.Cell(rowNumber, 3).Value = ratingNumber
                        wsFin.Cell(rowNumber, 4).Value = asofdate ' it should be in first
                        wsFin.Cell(rowNumber, 5).Value = Original_Face_ValueSum
                        wsFin.Cell(rowNumber, 6).Value = Current_Face_Sum
                        wsFin.Cell(rowNumber, 7).Value = accruedInterestSum
                        wsFin.Cell(rowNumber, 8).Value = MarketValueSum
                        wsFin.Cell(rowNumber, 9).Value = Stat_Book_ValueSum
                        wsFin.Cell(rowNumber, 10).Value = Stat_Carry_ValueSum
                        wsFin.Cell(rowNumber, 11).Value = YieldtoMaturitysum * MarketValueSum
                        wsFin.Cell(rowNumber, 12).Value = EffectiveDurationsum * MarketValueSum
                        wsFin.Cell(rowNumber, 13).Value = SpreadDurationsum * MarketValueSum
                        wsFin.Cell(rowNumber, 14).Value = OASsum * MarketValueSum
                        wsFin.Cell(rowNumber, 15).Value = DiscountMarginsum * MarketValueSum
                        wsFin.Cell(rowNumber, 16).Value = WeightedAvgLifesum * MarketValueSum
                        rowNumber += 1
                        TotalOriginal_Face_ValueSum += Original_Face_ValueSum
                        TotalCurrent_Face_Sum += Current_Face_Sum
                        TotalaccruedInterestSum += accruedInterestSum
                        TotalMarketValueSum += MarketValueSum
                        TotalStat_Book_ValueSum += Stat_Book_ValueSum
                        TotalStat_Carry_ValueSum += Stat_Carry_ValueSum
                        Console.WriteLine("Funding Type: " & fundingType & ", Category: " & assetCategory & ", Rating: " & ratingNumber & ", As Of Date: " & asofdate & ", Original_Face_ValueSum: " & Original_Face_ValueSum & ", Current_Face_Sum: " & Current_Face_Sum & ", accruedInterestSum: " & accruedInterestSum & ",MarketValueSum: " & MarketValueSum & ",Stat_Book_ValueSum: " & Stat_Book_ValueSum & ",Stat_Carry_ValueSum: " & Stat_Carry_ValueSum)

                    Next
                Next
            Next


            Console.WriteLine("TotalOriginal_Face_ValueSum = " & TotalOriginal_Face_ValueSum)
            Console.WriteLine("TotalCurrent_Face_Sum = " & TotalCurrent_Face_Sum)
            Console.WriteLine("TotalaccruedInterestSum = " & TotalaccruedInterestSum)
            Console.WriteLine("TotalMarketValueSum = " & TotalMarketValueSum)
            Console.WriteLine("TotalStat_Book_ValueSum = " & TotalStat_Book_ValueSum)
            Console.WriteLine("TotalStat_Carry_ValueSum = " & TotalStat_Carry_ValueSum)
            wsFin.Cell(rowNumber, 5).Value = "TotalOriginal_Face_ValueSum"
            wsFin.Cell(rowNumber, 5).Value = TotalOriginal_Face_ValueSum
            wsFin.Cell(rowNumber, 6).Value = "TotalCurrent_Face_Sum"
            wsFin.Cell(rowNumber, 6).Value = TotalCurrent_Face_Sum
            wsFin.Cell(rowNumber, 7).Value = "TotalaccruedInterestSum"
            wsFin.Cell(rowNumber, 7).Value = TotalaccruedInterestSum
            wsFin.Cell(rowNumber, 8).Value = "TotalMarketValueSum"
            wsFin.Cell(rowNumber, 8).Value = TotalMarketValueSum
            wsFin.Cell(rowNumber, 9).Value = "TotalStat_Book_ValueSum"
            wsFin.Cell(rowNumber, 9).Value = TotalStat_Book_ValueSum
            wsFin.Cell(rowNumber, 10).Value = "TotalStat_Carry_ValueSum"
            wsFin.Cell(rowNumber, 10).Value = TotalStat_Carry_ValueSum

            Dim TotalYieldtoMaturity As Double = 0
            Dim TotalEffectiveDuration As Double = 0
            Dim TotalSpreadDuration As Double = 0
            Dim TotalOAS As Double = 0
            Dim TotalDiscountMargin As Double = 0
            Dim TotalWeightedAvgLife As Double = 0

            rowNumber = 2
            Dim columns As Integer() = {11, 12, 13, 14, 15, 16}
            Dim totals As Double() = {TotalYieldtoMaturity, TotalEffectiveDuration, TotalSpreadDuration, TotalOAS, TotalDiscountMargin, TotalWeightedAvgLife}

            Dim loopcount = resultMap("SRRe_Asset_Funding").Count + resultMap("SRRe_Category").Count + resultMap("SRRe_Rating").Count

            'While loopcount > 0
            '    loopcount -= 1
            For Each fundingType As String In resultMap("SRRe_Asset_Funding")
                For Each assetCategory As String In resultMap("SRRe_Category")
                    For Each ratingNumber As String In resultMap("SRRe_Rating")
                        For i As Integer = 0 To columns.Length - 1
                            Dim col As Integer = columns(i)
                            Dim cellValue As Object = wsFin.Cell(rowNumber, col).Value
                            Dim numValue As Double = 0

                            If Double.TryParse(cellValue.ToString(), numValue) Then
                                Dim result As Double = 0

                                If TotalMarketValueSum <> 0 Then
                                    result = numValue / TotalMarketValueSum

                                    ' Avoid writing NaN or Infinity
                                    If Double.IsNaN(result) OrElse Double.IsInfinity(result) Then
                                        result = 0
                                    End If
                                Else
                                    result = 0
                                End If

                                wsFin.Cell(rowNumber, col).Value = result
                                totals(i) += numValue
                            Else
                                wsFin.Cell(rowNumber, col).Value = 0
                            End If
                        Next

                        rowNumber += 1
                    Next
                Next
            Next

            'End While


            ' Assign updated total values back
            TotalYieldtoMaturity = totals(0)
            TotalEffectiveDuration = totals(1)
            TotalSpreadDuration = totals(2)
            TotalOAS = totals(3)
            TotalDiscountMargin = totals(4)
            TotalWeightedAvgLife = totals(5)
            wsFin.Cell(rowNumber, 11).Value = "TotalYieldtoMaturity"
            wsFin.Cell(rowNumber, 11).Value = TotalYieldtoMaturity
            wsFin.Cell(rowNumber, 12).Value = "TotalEffectiveDuration"
            wsFin.Cell(rowNumber, 12).Value = TotalEffectiveDuration
            wsFin.Cell(rowNumber, 13).Value = "TotalSpreadDuration"
            wsFin.Cell(rowNumber, 13).Value = TotalSpreadDuration
            wsFin.Cell(rowNumber, 14).Value = "TotalOAS"
            wsFin.Cell(rowNumber, 14).Value = TotalOAS
            wsFin.Cell(rowNumber, 15).Value = "TotalDiscountMargin"
            wsFin.Cell(rowNumber, 15).Value = TotalDiscountMargin
            wsFin.Cell(rowNumber, 16).Value = "TotalWeightedAvgLife"
            wsFin.Cell(rowNumber, 16).Value = TotalWeightedAvgLife


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

                Console.WriteLine(newFileName) ' Output: D:\BMA_TP11_30_2024.xlsx
            Else
                Console.WriteLine("Invalid date format.")
            End If

            ' Combine directory path with the new filename
            'Dim newFilePath As String = Path.Combine(directoryPath, newFileName)

            Console.WriteLine(newFileName)

            wbFin.SaveAs(newFileName)
            Console.WriteLine("Report generated at: " & outputPath)
        End Using

    End Sub

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
