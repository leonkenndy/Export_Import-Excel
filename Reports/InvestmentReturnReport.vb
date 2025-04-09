Imports System.Data.OleDb
Imports ClosedXML.Excel
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Globalization
Imports DocumentFormat.OpenXml.Spreadsheet
Imports System.ComponentModel
Public Class InvestmentReturnReport

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
    Sub GenerateInvestmentReturnReport()
        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'InvestmentReturnReport
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
            wsFin.Cell(1, 2).Value = "Coupon Type"
            wsFin.Cell(1, 3).Value = "Category"
            wsFin.Cell(1, 4).Value = "Rating"
            wsFin.Cell(1, 5).Value = "As Of Date"
            wsFin.Cell(1, 6).Value = "Accrued Interest Market Value Sum"
            ' Queries to select distinct values from different columns
            Dim SRRe_Asset_Fundingquery As String = "SELECT DISTINCT [SRRe Asset Funding] FROM [" & tableName & "]"
            Dim Coupon_Type_query As String = "SELECT DISTINCT [SRRe Coupon Type] FROM [" & tableName & "]"
            Dim SRRe_Category_query As String = "SELECT DISTINCT [SRRe Asset Category] FROM [" & tableName & "]"
            Dim SRRe_Rating_query As String = "SELECT DISTINCT [SRRe Rating Number] FROM [" & tableName & "]"

            ' Create a list of lists to store the results
            Dim resultMap As New Dictionary(Of String, List(Of String))()

            ' Create a connection to the database


            Try

                ' Run each query and store the results in the dictionary
                resultMap("SRRe_Asset_Funding") = ExecuteQuery(conn, SRRe_Asset_Fundingquery)
                resultMap("Coupon_Type") = ExecuteQuery(conn, Coupon_Type_query)
                resultMap("SRRe_Rating") = ExecuteQuery(conn, SRRe_Rating_query)
                resultMap("SRRe_Category") = ExecuteQuery(conn, SRRe_Category_query)

            Catch ex As Exception
                Console.WriteLine("Error: " & ex.Message)
            End Try
            Dim TotalMarketValueSum As Long = 0
            Dim rowNumber As Integer = 2
            ' Example loop through each combination
            For Each fundingType As String In resultMap("SRRe_Asset_Funding")
                For Each couponType As String In resultMap("Coupon_Type")
                    For Each assetCategory As String In resultMap("SRRe_Category")
                        For Each ratingNumber As String In resultMap("SRRe_Rating")

                            ' Filter the data with specific conditions
                            Dim accruedInterestMarketValueSum As Decimal = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Asset Funding") = fundingType AndAlso
                                                                                                 r.Field(Of String)("Coupon Type") = couponType AndAlso
                                                                                                 r.Field(Of String)("SRRe Asset Category") = assetCategory AndAlso
                                                                                                 r.Field(Of String)("SRRe Rating Number") = ratingNumber AndAlso
                                                                                                 r.Field(Of String)("As Of Date") = asofdate
                                                                                                 ).Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Market Value"))))

                            ' Output the result for the current combination
                            ' Write the data to the next row
                            wsFin.Cell(rowNumber, 1).Value = fundingType
                            wsFin.Cell(rowNumber, 2).Value = couponType
                            wsFin.Cell(rowNumber, 3).Value = assetCategory
                            wsFin.Cell(rowNumber, 4).Value = ratingNumber
                            wsFin.Cell(rowNumber, 5).Value = asofdate
                            wsFin.Cell(rowNumber, 6).Value = accruedInterestMarketValueSum
                            rowNumber += 1
                            TotalMarketValueSum += accruedInterestMarketValueSum
                            Console.WriteLine("Funding Type: " & fundingType & ", Coupon Type: " & couponType & ", Category: " & assetCategory & ", Rating: " & ratingNumber & ", As Of Date: " & asofdate)
                            Console.WriteLine("Accrued Interest Market Value Sum = " & accruedInterestMarketValueSum)
                        Next
                    Next
                Next
            Next
            wsFin.Cell(rowNumber, 6).Value = "TotalMarketValueSum"
            wsFin.Cell(rowNumber, 6).Value = TotalMarketValueSum

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
                newFileName = directoryPath & "\InvestmentReturn" & formattedDate & ".xlsx"

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
