
Imports ClosedXML.Excel
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Globalization

Public Class BMA_TPReport
    Private dbPath As String
    Private tableName As String
    Private outputPath As String
    Private asofdate As String
    Private rowIndex As Integer
    Dim wb As New XLWorkbook()
    Dim ws = wb.Worksheets.Add("Report")
    Public Sub New(ByVal dbPath As String, ByVal tableName As String, ByVal outputPath As String, ByVal asofdate As String)
        Me.dbPath = dbPath
        Me.tableName = tableName
        Me.outputPath = outputPath
        Me.asofdate = asofdate
        Me.rowIndex = 0

    End Sub

    Public Sub GenerateReport()
        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'BMA TP Report 

        ' Define the list of SRRe Asset Types
        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
        Dim dataTable As New DataTable()
        Dim resultMap As New Dictionary(Of String, List(Of String))()
        Using conn As New OleDbConnection(connectionString)
            conn.Open()

            ' Query all required data
            Dim query As String = "SELECT * FROM [" & tableName & "] WHERE [As of Date] = ?"
            Using cmd As New OleDbCommand(query, conn)
                ' Add parameter to avoid SQL injection and formatting issues
                cmd.Parameters.AddWithValue("@asofdate", asofdate)
                Dim SRRe_Asset_Typequery As String = "SELECT DISTINCT [SRRe Asset Type] FROM [" & tableName & "]"
                Dim SRRe_IG_query As String = "SELECT DISTINCT [SRRe IG] FROM [" & tableName & "]"
                Dim SRRe_Port_Code_query As String = "SELECT DISTINCT [SRRe Port Code] FROM [" & tableName & "]"

                Try

                    ' Run each query and store the results in the dictionary
                    resultMap("SRRe_Asset_Typequery") = ExecuteQuery(conn, SRRe_Asset_Typequery)
                    resultMap("SRRe_IG_query") = ExecuteQuery(conn, SRRe_IG_query)
                    resultMap("SRRe_Port_Code_query") = ExecuteQuery(conn, SRRe_Port_Code_query)

                Catch ex As Exception
                    Console.WriteLine("Error: " & ex.Message)
                End Try
                Using reader As OleDbDataReader = cmd.ExecuteReader()
                    dataTable.Load(reader)
                End Using
            End Using
        End Using



        ' Example Header
        ws.Cell(1, 1).Value = "Summary Report BMA TP"
        ws.Cell(2, 1).Value = "Generated On"
        ws.Cell(2, 2).Value = DateTime.Now.ToString("yyyy-MM-dd")
        ws.Cell(2, 3).Value = "ME Holdings As of Date"
        ws.Cell(2, 4).Value = asofdate
        ' Example Calculations
        rowIndex = 4
        ws.Cell(rowIndex, 1).Value = "SRRe Port Code"
        ws.Cell(rowIndex, 2).Value = "SRRe Asset Type"
        ws.Cell(rowIndex, 3).Value = "IG"
        ws.Cell(rowIndex, 4).Value = "Non-IG"
        rowIndex += 1

        For Each PortCodeType As String In resultMap("SRRe_Port_Code_query")
            For Each Asset_Type As String In resultMap("SRRe_Asset_Typequery")

                'Debugging
                Console.WriteLine("Funding Type: " & Asset_Type & ", Rating Number: " & PortCodeType)

                Dim accruedInterestMarketValueSumIG As Decimal = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Asset Funding") = "Funded" AndAlso
                                                                  r.Field(Of String)("SRRe Port Code") = PortCodeType AndAlso
                                                                  r.Field(Of String)("SRRe Asset Type") = Asset_Type AndAlso
                                                                  r.Field(Of String)("SRRe IG") = "IG" AndAlso
                                                                  r.Field(Of String)("As Of Date") = asofdate).
                                                                        Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Accrued Interest"))) +
                                                               Convert.ToDecimal(Convert.ToDouble(r("Market Value"))))

                Dim accruedInterestMarketValueSum As Decimal = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Asset Funding") = "Funded" AndAlso
                                                              r.Field(Of String)("SRRe Port Code") = PortCodeType AndAlso
                                                              r.Field(Of String)("SRRe Asset Type") = Asset_Type AndAlso
                                                              r.Field(Of String)("SRRe IG") <> "IG" AndAlso
                                                              r.Field(Of String)("As Of Date") = asofdate).
                                                            Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Accrued Interest"))) +
                                                                           Convert.ToDecimal(Convert.ToDouble(r("Market Value"))))


                ws.Cell(rowIndex, 1).Value = PortCodeType
                ws.Cell(rowIndex, 2).Value = Asset_Type
                ws.Cell(rowIndex, 3).Value = accruedInterestMarketValueSumIG
                ws.Cell(rowIndex, 4).Value = accruedInterestMarketValueSum

                rowIndex += 1

            Next
        Next

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
            newFileName = directoryPath & "\BMA_TP " & formattedDate & ".xlsx"

            Console.WriteLine(newFileName)
        Else
            Console.WriteLine("Invalid date format.")
        End If

        Console.WriteLine(newFileName)

        wb.SaveAs(newFileName)
        Console.WriteLine("Report generated at: " & outputPath)
    End Sub

    ' Function to execute a query and return the results as a List(Of String)
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
