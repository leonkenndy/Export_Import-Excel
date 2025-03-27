Imports System.Data.OleDb
Imports ClosedXML.Excel
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Globalization
Public Class ReportGenerator
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
        End Using
        ' Create Excel
        'Dim wb As New XLWorkbook()
        'Dim ws = wb.Worksheets.Add("Report")

        ' Example Header
        ws.Cell(1, 1).Value = "Summary Report"
        ws.Cell(2, 1).Value = "Generated On"
        ws.Cell(2, 2).Value = DateTime.Now.ToString("yyyy-MM-dd")
        ws.Cell(2, 3).Value = "ME Holdings As of Date"
        ws.Cell(2, 4).Value = asofdate
        ' Example Calculations
        rowIndex = 4
        ' Total Investment-Grade Corporate

        Dim igCorpSum = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                             r.Field(Of String)("SRRe Asset Type") = "Corp Sec Public").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Investment-Grade Corporate"
        ws.Cell(rowIndex, 2).Value = igCorpSum
        rowIndex += 1


        ' IG Corp Public
        Dim igCorpPublicSum = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                             r.Field(Of String)("SRRe Asset Type") = "Corp Sec Public" AndAlso
                                                             r.Field(Of String)("SRRe IG") = "IG").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "IG Corp Public"
        ws.Cell(rowIndex, 2).Value = igCorpPublicSum
        rowIndex += 1

        ' IG Corp Non-Public
        Dim igCorpNonPublicSum = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                                r.Field(Of String)("SRRe Asset Type") = "Corp Sec Private" AndAlso
                                                                r.Field(Of String)("SRRe IG") = "IG").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "IG Corp Non Public"
        ws.Cell(rowIndex, 2).Value = igCorpNonPublicSum
        rowIndex += 1

        ' Total Below Investment-Grade Corporates

        Dim bigCorpSum = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                             r.Field(Of String)("SRRe Asset Type") = "Corp Sec Private").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Below Investment-Grade Corporate"
        ws.Cell(rowIndex, 2).Value = bigCorpSum
        rowIndex += 1

        ' IG Corp Non-Public
        Dim nonIGCorpPublicSum = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                                r.Field(Of String)("SRRe Asset Type") = "Corp Sec Private" AndAlso
                                                                r.Field(Of String)("SRRe IG") = "BIG").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Non-IG Corp Public"
        ws.Cell(rowIndex, 2).Value = nonIGCorpPublicSum
        rowIndex += 1

        Dim nonIGCorpNonPublicSum = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                                r.Field(Of String)("SRRe Asset Type") = "Corp Sec Private" AndAlso
                                                                r.Field(Of String)("SRRe IG") = "BIG").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Non-IG Corp Non Public"
        ws.Cell(rowIndex, 2).Value = nonIGCorpNonPublicSum
        rowIndex += 1
        '------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Total CLOs

        Dim CLOsSum = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                             r.Field(Of String)("SRRe Asset Type") = "CLO").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "CLOs"
        ws.Cell(rowIndex, 2).Value = CLOsSum
        rowIndex += 1


        Dim IGCLOs = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                                r.Field(Of String)("SRRe Asset Type") = "CLO" AndAlso
                                                                r.Field(Of String)("SRRe IG") = "IG").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "IG CLOs"
        ws.Cell(rowIndex, 2).Value = IGCLOs
        rowIndex += 1

        Dim NonIGCLOs = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                                r.Field(Of String)("SRRe Asset Type") = "CLO" AndAlso
                                                                r.Field(Of String)("SRRe IG") = "IG").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Non-IG CLOs"
        ws.Cell(rowIndex, 2).Value = NonIGCLOs
        rowIndex += 1

        '------------------------------------------------------------------------------------------------------------------------------------------------------------------

        'Total Collateral Loans
        Dim totalcollateral_loansSum = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                                r.Field(Of String)("SRRe Asset Type") = "Collateral Loans" Or
                                                                r.Field(Of String)("SRRe Asset Type") = "Syndicated Bank Loans").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Total Collateral Loans"
        ws.Cell(rowIndex, 2).Value = totalcollateral_loansSum
        rowIndex += 1

        Dim collateral_loans = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                                r.Field(Of String)("SRRe Asset Type") = "Collateral Loans").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Collateral Loans"
        ws.Cell(rowIndex, 2).Value = collateral_loans
        rowIndex += 1

        Dim syndicated_bank_loans = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                                r.Field(Of String)("SRRe Asset Type") = "Syndicated Bank Loans").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Syndicated Bank Loans"
        ws.Cell(rowIndex, 2).Value = syndicated_bank_loans
        rowIndex += 1

        '------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'CML

        Dim totalcmls = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                                r.Field(Of String)("SRRe Asset Type") = "Commercial Mortgage Loans").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "IG Commercial Mortgage Loans"
        ws.Cell(rowIndex, 2).Value = totalcmls
        rowIndex += 1

        Dim igcmls = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                                r.Field(Of String)("SRRe Asset Type") = "Commercial Mortgage Loans" AndAlso
                                                                r.Field(Of String)("SRRe IG") = "IG").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "IG Commercial Mortgage Loans"
        ws.Cell(rowIndex, 2).Value = igcmls
        rowIndex += 1

        Dim nonigcml = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                                  r.Field(Of String)("SRRe Asset Type") = "Commercial Mortgage Loans" AndAlso
                                                                  r.Field(Of String)("SRRe IG") = "BIG").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Non-IG Commercial Mortgage Loans"
        ws.Cell(rowIndex, 2).Value = nonigcml
        rowIndex += 1

        '------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'ABS CMBS RMBS
        'Others 
        Dim othersStructureAssets = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                               (r.Field(Of String)("SRRe Asset Type") = "ABS" Or
                                r.Field(Of String)("SRRe Asset Type") = "CMBS" Or
                                r.Field(Of String)("SRRe Asset Type") = "RMBS")).Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Others Sturctured Assets"
        ws.Cell(rowIndex, 2).Value = othersStructureAssets
        rowIndex += 1

        Dim abs = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                             r.Field(Of String)("SRRe Asset Type") = "ABS").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "ABS"
        ws.Cell(rowIndex, 2).Value = abs
        rowIndex += 1

        Dim cmbs = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                              r.Field(Of String)("SRRe Asset Type") = "CMBS").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "CMBS"
        ws.Cell(rowIndex, 2).Value = cmbs
        rowIndex += 1

        Dim rmbs = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                              r.Field(Of String)("SRRe Asset Type") = "RMBS").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "RMBS"
        ws.Cell(rowIndex, 2).Value = rmbs
        rowIndex += 1

        '------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'Direct investments and Preferred Stock
        Dim others = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                             r.Field(Of String)("SRRe Asset Type") = "Direct Investments" Or
                                                              r.Field(Of String)("SRRe Asset Type") = "Preferred Stock").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Others"
        ws.Cell(rowIndex, 2).Value = others
        rowIndex += 1


        Dim directInvestments = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                             r.Field(Of String)("SRRe Asset Type") = "Direct Investments").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Direct Investments"
        ws.Cell(rowIndex, 2).Value = directInvestments
        rowIndex += 1

        Dim preferredStock = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                              r.Field(Of String)("SRRe Asset Type") = "Preferred Stock").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Preferred Stock"
        ws.Cell(rowIndex, 2).Value = preferredStock
        rowIndex += 1
        '------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'Commercial Paper
        Dim commercialPaper = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                             r.Field(Of String)("SRRe Asset Type") = "Commercial Paper").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Commercial Paper"
        ws.Cell(rowIndex, 2).Value = commercialPaper
        rowIndex += 1
        'TODO Need to know answers 
        'Cash and Equivalents - Pre-QSSA
        Dim cashEquivalentPreQssa = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                             r.Field(Of String)("SRRe Asset Type") = "Commercial Paper").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Cash and Equivalents - Pre-QSSA(IGNORE)"
        ws.Cell(rowIndex, 2).Value = cashEquivalentPreQssa
        rowIndex += 1

        'Cash and Equivalents
        Dim cashEquivalent = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                             r.Field(Of String)("SRRe Asset Type") = "Commercial Paper").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Cash and Equivalents(IGNORE)"
        ws.Cell(rowIndex, 2).Value = cashEquivalent
        rowIndex += 1

        '    ' Save
        '    wb.SaveAs(outputPath)
        '    Console.WriteLine("Report generated at: " & outputPath)
        'End Sub



        'Sub GenerateReportOC()
        '    Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
        '    Dim dataTable As New DataTable()

        '    Using conn As New OleDbConnection(connectionString)
        '        conn.Open()

        '        ' Query all required data
        '        Dim query As String = "SELECT * FROM [" & tableName & "] WHERE [As of Date] = ?"

        '        Using cmd As New OleDbCommand(query, conn)
        '            ' Add parameter to avoid SQL injection and formatting issues
        '            cmd.Parameters.AddWithValue("@asofdate", asofdate)

        '            Using reader As OleDbDataReader = cmd.ExecuteReader()
        '                dataTable.Load(reader)
        '            End Using
        '        End Using
        '    End Using
        ' Create Excel


        '' Example Header
        'ws.Cell(1, 1).Value = "Summary Report"
        'ws.Cell(2, 1).Value = "Generated On"
        'ws.Cell(2, 2).Value = DateTime.Now.ToString("yyyy-MM-dd")
        'ws.Cell(2, 3).Value = "ME Holdings As of Date"
        'ws.Cell(2, 4).Value = asofdate
        ' Example Calculations
        'Dim rowIndex As Integer = 4
        ' Total Investment-Grade Corporate

        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'OC Account
        ' IG Corp Public
        ' Total Investment-Grade Corporate
        ws.Cell(rowIndex, 1).Value = "OC Account"
        Dim igCorpSumOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                       r.Field(Of String)("SRRe Asset Type") = "Corp Sec Public").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Investment-Grade Corporate"
        ws.Cell(rowIndex, 2).Value = igCorpSumOC
        rowIndex += 1


        Dim igCorpPublicSumOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                             r.Field(Of String)("SRRe Asset Type") = "Corp Sec Public" AndAlso
                                                             r.Field(Of String)("SRRe IG") = "IG").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "IG Corp Public"
        ws.Cell(rowIndex, 2).Value = igCorpPublicSumOC
        rowIndex += 1

        ' Total Below Investment-Grade Corporates

        Dim bigCorpSumOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                         r.Field(Of String)("SRRe Asset Type") = "Corp Sec Private").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Below Investment-Grade Corporate"
        ws.Cell(rowIndex, 2).Value = bigCorpSumOC
        rowIndex += 1

        ' IG Corp Non-Public
        Dim igCorpNonPublicSumOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                                r.Field(Of String)("SRRe Asset Type") = "Corp Sec Private" AndAlso
                                                                r.Field(Of String)("SRRe IG") = "IG").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "IG Corp Non Public"
        ws.Cell(rowIndex, 2).Value = igCorpNonPublicSumOC
        rowIndex += 1


        ' IG Corp Non-Public
        Dim nonIGCorpPublicSumOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                                r.Field(Of String)("SRRe Asset Type") = "Corp Sec Private" AndAlso
                                                                r.Field(Of String)("SRRe IG") = "BIG").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Non-IG Corp Public"
        ws.Cell(rowIndex, 2).Value = nonIGCorpPublicSumOC
        rowIndex += 1

        Dim nonIGCorpNonPublicSumOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                                r.Field(Of String)("SRRe Asset Type") = "Corp Sec Private" AndAlso
                                                                r.Field(Of String)("SRRe IG") = "BIG").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Non-IG Corp Non Public"
        ws.Cell(rowIndex, 2).Value = nonIGCorpNonPublicSumOC
        rowIndex += 1
        '------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Total CLOs

        Dim CLOsSumOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                             r.Field(Of String)("SRRe Asset Type") = "CLO").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "CLOs"
        ws.Cell(rowIndex, 2).Value = CLOsSumOC
        rowIndex += 1


        Dim IGCLOsOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                                r.Field(Of String)("SRRe Asset Type") = "CLO" AndAlso
                                                                r.Field(Of String)("SRRe IG") = "IG").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "IG CLOs"
        ws.Cell(rowIndex, 2).Value = IGCLOsOC
        rowIndex += 1

        Dim NonIGCLOsOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                                r.Field(Of String)("SRRe Asset Type") = "CLO" AndAlso
                                                                r.Field(Of String)("SRRe IG") = "BIG").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Non-IG CLOs"
        ws.Cell(rowIndex, 2).Value = NonIGCLOsOC
        rowIndex += 1

        '------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'Total Collateral Loans
        Dim totalcollateral_loansSumOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                            r.Field(Of String)("SRRe Asset Type") = "Collateral Loans" Or
                                                            r.Field(Of String)("SRRe Asset Type") = "Syndicated Bank Loans").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Total Collateral Loans"
        ws.Cell(rowIndex, 2).Value = totalcollateral_loansSumOC
        rowIndex += 1

        Dim collateral_loansOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                                r.Field(Of String)("SRRe Asset Type") = "Collateral Loans").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Collateral Loans"
        ws.Cell(rowIndex, 2).Value = collateral_loansOC
        rowIndex += 1

        Dim syndicated_bank_loansOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                                r.Field(Of String)("SRRe Asset Type") = "Syndicated Bank Loans").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Syndicated Bank Loans"
        ws.Cell(rowIndex, 2).Value = syndicated_bank_loansOC
        rowIndex += 1

        '------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'CML

        Dim totalcmlsOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                            r.Field(Of String)("SRRe Asset Type") = "Commercial Mortgage Loans").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "IG Commercial Mortgage Loans"
        ws.Cell(rowIndex, 2).Value = totalcmlsOC
        rowIndex += 1


        Dim igcmlsOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                                r.Field(Of String)("SRRe Asset Type") = "Commercial Mortgage Loans" AndAlso
                                                                r.Field(Of String)("SRRe IG") = "IG").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "IG Commercial Mortgage Loans"
        ws.Cell(rowIndex, 2).Value = igcmlsOC
        rowIndex += 1

        Dim nonigcmlOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                                  r.Field(Of String)("SRRe Asset Type") = "Commercial Mortgage Loans" AndAlso
                                                                  r.Field(Of String)("SRRe IG") = "BIG").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Non-IG Commercial Mortgage Loans"
        ws.Cell(rowIndex, 2).Value = nonigcmlOC
        rowIndex += 1

        '------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'ABS CMBS RMBS
        'Others 
        Dim othersStructureAssetsOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                               (r.Field(Of String)("SRRe Asset Type") = "ABS" Or
                                r.Field(Of String)("SRRe Asset Type") = "CMBS" Or
                                r.Field(Of String)("SRRe Asset Type") = "RMBS")).Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Others Sturctured Assets"
        ws.Cell(rowIndex, 2).Value = othersStructureAssetsOC
        rowIndex += 1

        Dim absOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                             r.Field(Of String)("SRRe Asset Type") = "ABS").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "ABS"
        ws.Cell(rowIndex, 2).Value = absOC
        rowIndex += 1

        Dim cmbsOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                              r.Field(Of String)("SRRe Asset Type") = "CMBS").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "CMBS"
        ws.Cell(rowIndex, 2).Value = cmbsOC
        rowIndex += 1

        Dim rmbsOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                              r.Field(Of String)("SRRe Asset Type") = "RMBS").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "RMBS"
        ws.Cell(rowIndex, 2).Value = rmbsOC
        rowIndex += 1

        '------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'Direct investments and Preferred Stock
        'Direct investments and Preferred Stock
        Dim othersOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                          r.Field(Of String)("SRRe Asset Type") = "Direct Investments" Or
                                                           r.Field(Of String)("SRRe Asset Type") = "Preferred Stock").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Others"
        ws.Cell(rowIndex, 2).Value = othersOC
        rowIndex += 1

        Dim directInvestmentsOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                             r.Field(Of String)("SRRe Asset Type") = "Direct Investments").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Direct Investments"
        ws.Cell(rowIndex, 2).Value = directInvestmentsOC
        rowIndex += 1

        Dim preferredStockOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                              r.Field(Of String)("SRRe Asset Type") = "Preferred Stock").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Preferred Stock"
        ws.Cell(rowIndex, 2).Value = preferredStockOC
        rowIndex += 1
        '------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'Commercial Paper
        Dim commercialPaperOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                             r.Field(Of String)("SRRe Asset Type") = "Commercial Paper").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Commercial Paper"
        ws.Cell(rowIndex, 2).Value = commercialPaperOC
        rowIndex += 1


        'TODO Need to know answers 
        'Cash and Equivalents - Pre-QSSA
        Dim cashEquivalentPreQssaOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                        r.Field(Of String)("SRRe Asset Type") = "Commercial Paper").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Cash and Equivalents - Pre-QSSA"
        ws.Cell(rowIndex, 2).Value = cashEquivalentPreQssa
        rowIndex += 1

        'Cash and Equivalents
        Dim cashEquivalentOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                        r.Field(Of String)("SRRe Asset Type") = "Commercial Paper").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Cash and Equivalents"
        ws.Cell(rowIndex, 2).Value = cashEquivalentOC
        rowIndex += 1
        ' Save





        '------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' FWH Plus OC Totals 
        ws.Cell(++rowIndex, 1).Value = "Totals FWH + OC"
        rowIndex += 1

        Dim totalCashAndEquivalents As Integer = cashEquivalent + cashEquivalentOC
        ws.Cell(rowIndex, 1).Value = "Cash and Equivalents"
        ws.Cell(rowIndex, 2).Value = totalCashAndEquivalents
        rowIndex += 1

        Dim totalCashEquivalentPreQssa As Integer = cashEquivalentPreQssa + cashEquivalentPreQssaOC
        ws.Cell(rowIndex, 1).Value = "Cash Equivalent Pre-QSSA"
        ws.Cell(rowIndex, 2).Value = totalCashEquivalentPreQssa
        rowIndex += 1

        Dim totalCommercialPaper As Integer = commercialPaper + commercialPaperOC
        ws.Cell(rowIndex, 1).Value = "Commercial Paper"
        ws.Cell(rowIndex, 2).Value = totalCommercialPaper
        rowIndex += 1

        Dim totalDirectInvestments As Integer = directInvestments + directInvestmentsOC
        ws.Cell(rowIndex, 1).Value = "Direct Investments"
        ws.Cell(rowIndex, 2).Value = totalDirectInvestments
        rowIndex += 1

        Dim totalIgCorpPublicSum As Integer = igCorpPublicSum + igCorpPublicSumOC
        ws.Cell(rowIndex, 1).Value = "IG Corporate Public Sum"
        ws.Cell(rowIndex, 2).Value = totalIgCorpPublicSum
        rowIndex += 1

        Dim totalIgCorpSum As Integer = igCorpSum + igCorpSumOC
        ws.Cell(rowIndex, 1).Value = "IG Corporate Sum"
        ws.Cell(rowIndex, 2).Value = totalIgCorpSum
        rowIndex += 1

        Dim totalIgCorpNonPublicSum As Integer = igCorpNonPublicSum + igCorpNonPublicSumOC
        ws.Cell(rowIndex, 1).Value = "IG Corporate Non-Public Sum"
        ws.Cell(rowIndex, 2).Value = totalIgCorpNonPublicSum
        rowIndex += 1

        Dim totalBigCorpSum As Integer = bigCorpSum + bigCorpSumOC
        ws.Cell(rowIndex, 1).Value = "Big Corporate Sum"
        ws.Cell(rowIndex, 2).Value = totalBigCorpSum
        rowIndex += 1

        Dim totalNonIgCml As Integer = nonigcml + nonigcmlOC
        ws.Cell(rowIndex, 1).Value = "Non-IG CML"
        ws.Cell(rowIndex, 2).Value = totalNonIgCml
        rowIndex += 1

        Dim totalNonIgCorpNonPublicSum As Integer = nonIGCorpNonPublicSum + nonIGCorpNonPublicSumOC
        ws.Cell(rowIndex, 1).Value = "Non-IG Corporate Non-Public Sum"
        ws.Cell(rowIndex, 2).Value = totalNonIgCorpNonPublicSum
        rowIndex += 1

        Dim totalNonIgCorpPublicSum As Integer = nonIGCorpPublicSum + nonIGCorpPublicSumOC
        ws.Cell(rowIndex, 1).Value = "Non-IG Corporate Public Sum"
        ws.Cell(rowIndex, 2).Value = totalNonIgCorpPublicSum
        rowIndex += 1

        Dim totalNonIgCLOs As Integer = NonIGCLOs + NonIGCLOsOC
        ws.Cell(rowIndex, 1).Value = "Non-IG CLOs"
        ws.Cell(rowIndex, 2).Value = totalNonIgCLOs
        rowIndex += 1

        Dim totalCollateralLoans As Integer = collateral_loans + collateral_loansOC
        ws.Cell(rowIndex, 1).Value = "Collateral Loans"
        ws.Cell(rowIndex, 2).Value = totalCollateralLoans
        rowIndex += 1

        Dim totalSyndicatedBankLoans As Integer = syndicated_bank_loans + syndicated_bank_loansOC
        ws.Cell(rowIndex, 1).Value = "Syndicated Bank Loans"
        ws.Cell(rowIndex, 2).Value = totalSyndicatedBankLoans
        rowIndex += 1

        Dim totalIgCmls As Integer = igcmls + igcmlsOC
        ws.Cell(rowIndex, 1).Value = "IG CMLs"
        ws.Cell(rowIndex, 2).Value = totalIgCmls
        rowIndex += 1

        Dim totalCmlssum As Integer = totalcmls + totalcmlsOC
        ws.Cell(rowIndex, 1).Value = "Total CMLs"
        ws.Cell(rowIndex, 2).Value = totalCmlssum
        rowIndex += 1

        Dim totalOtherStructuredAssets As Integer = othersStructureAssets + othersStructureAssetsOC
        ws.Cell(rowIndex, 1).Value = "Other Structured Assets"
        ws.Cell(rowIndex, 2).Value = totalOtherStructuredAssets
        rowIndex += 1

        Dim totalOthers As Integer = others + othersOC
        ws.Cell(rowIndex, 1).Value = "Others"
        ws.Cell(rowIndex, 2).Value = totalOthers
        rowIndex += 1

        Dim totalPreferredStock As Integer = preferredStock + preferredStockOC
        ws.Cell(rowIndex, 1).Value = "Preferred Stock"
        ws.Cell(rowIndex, 2).Value = totalPreferredStock
        rowIndex += 1
        '------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Accrued Interest 

        Dim Accrued_InterestFWH = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                             r.Field(Of String)("As Of Date") = asofdate).Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Accrued Interest"))))

        ws.Cell(rowIndex, 1).Value = "Accrued Interest FWH"
        ws.Cell(rowIndex, 2).Value = Accrued_InterestFWH
        rowIndex += 1

        Dim Accrued_InterestOC = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                                                             r.Field(Of String)("As Of Date") = asofdate).Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Accrued Interest"))))

        ws.Cell(rowIndex, 1).Value = "Accrued Interest OC"
        ws.Cell(rowIndex, 2).Value = Accrued_InterestOC
        rowIndex += 1

        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'Total Percentage
        ' Function to calculate and insert percentage in Excel
        ' Calculate Grand Total
        Dim grandTotal As Integer = totalCashAndEquivalents + totalCashEquivalentPreQssa + totalCommercialPaper +
                            totalDirectInvestments + totalIgCorpPublicSum + totalIgCorpSum +
                            totalIgCorpNonPublicSum + totalBigCorpSum + totalNonIgCml +
                            totalNonIgCorpNonPublicSum + totalNonIgCorpPublicSum + totalNonIgCLOs +
                            totalCollateralLoans + totalSyndicatedBankLoans + totalIgCmls +
                            totalCmlssum + totalOtherStructuredAssets + totalOthers + totalPreferredStock

        ' Store Grand Total in Excel
        ws.Cell(rowIndex, 1).Value = "Grand Total"
        ws.Cell(rowIndex, 2).Value = grandTotal
        rowIndex += 1

        ws.Cell(rowIndex, 1).Value = "Total %"
        rowIndex += 1
        ' Insert percentages for each category
        InsertPercentage("Cash and Equivalents", totalCashAndEquivalents, grandTotal)
        InsertPercentage("Cash Equivalent Pre-QSSA", totalCashEquivalentPreQssa, grandTotal)
        InsertPercentage("Commercial Paper", totalCommercialPaper, grandTotal)
        InsertPercentage("Direct Investments", totalDirectInvestments, grandTotal)
        InsertPercentage("IG Corporate Public Sum", totalIgCorpPublicSum, grandTotal)
        InsertPercentage("IG Corporate Sum", totalIgCorpSum, grandTotal)
        InsertPercentage("IG Corporate Non-Public Sum", totalIgCorpNonPublicSum, grandTotal)
        InsertPercentage("Big Corporate Sum", totalBigCorpSum, grandTotal)
        InsertPercentage("Non-IG CML", totalNonIgCml, grandTotal)
        InsertPercentage("Non-IG Corporate Non-Public Sum", totalNonIgCorpNonPublicSum, grandTotal)
        InsertPercentage("Non-IG Corporate Public Sum", totalNonIgCorpPublicSum, grandTotal)
        InsertPercentage("Non-IG CLOs", totalNonIgCLOs, grandTotal)
        InsertPercentage("Collateral Loans", totalCollateralLoans, grandTotal)
        InsertPercentage("Syndicated Bank Loans", totalSyndicatedBankLoans, grandTotal)
        InsertPercentage("IG CMLs", totalIgCmls, grandTotal)
        InsertPercentage("Total CMLs", totalCmlssum, grandTotal)
        InsertPercentage("Other Structured Assets", totalOtherStructuredAssets, grandTotal)
        InsertPercentage("Others", totalOthers, grandTotal)
        InsertPercentage("Preferred Stock", totalPreferredStock, grandTotal)
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
            newFileName = "D:\ME_Holdings " & formattedDate & ".xlsx"

            Console.WriteLine(newFileName) ' Output: D:\BMA_TP11_30_2024.xlsx
        Else
            Console.WriteLine("Invalid date format.")
        End If


        wb.SaveAs(newFileName)
        Console.WriteLine("Report generated at: " & newFileName)

    End Sub

    Sub GenerateBMA_TPReport()
        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'BMA TP Report 

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
        End Using

        Dim wbBMA As New XLWorkbook()
        Dim wsBMA = wbBMA.Worksheets.Add("Report BMA TP")

        ' Example Header
        wsBMA.Cell(1, 1).Value = "Summary Report BMA TP"
        wsBMA.Cell(2, 1).Value = "Generated On"
        wsBMA.Cell(2, 2).Value = DateTime.Now.ToString("yyyy-MM-dd")
        wsBMA.Cell(2, 3).Value = "ME Holdings As of Date"
        wsBMA.Cell(2, 4).Value = asofdate
        ' Example Calculations
        Dim rowIndexBMA = 4
        ' Total Investment-Grade Corporate
        Dim srreAssetTypes As New List(Of String) From {
                                                            "ABS",
                                                            "Bank Deposit",
                                                            "CLO",
                                                            "CMBS",
                                                            "Collateral Loans",
                                                            "Commercial Mortgage Loans",
                                                            "Commercial Paper",
                                                            "Corp Sec 144a",
                                                            "Corp Sec Private",
                                                            "Corp Sec Public",
                                                            "Syndicated Bank Loans"
                                                        }

        ' Loop through each asset type 
        'FWH 'IG
        For Each assetType As String In srreAssetTypes
            ' Calculate Accrued Interest and Market Value Sum for the current asset type
            Dim accruedInterestMarketValueSum As Decimal = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Asset Funding") = "Funded" AndAlso
                          r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                          r.Field(Of String)("SRRe Asset Type") = assetType AndAlso
                          r.Field(Of String)("SRRe IG") = "IG" AndAlso
                          r.Field(Of String)("As Of Date") = asofdate).
        Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Accrued Interest"))) +
                       Convert.ToDecimal(Convert.ToDouble(r("Market Value"))))

            ' Write results to Excel
            wsBMA.Cell(rowIndexBMA, 1).Value = "Accrued Interest & Market Value FWH - " & assetType
            wsBMA.Cell(rowIndexBMA, 2).Value = accruedInterestMarketValueSum
            rowIndexBMA += 1
        Next
        'BIG
        For Each assetType As String In srreAssetTypes
            ' Calculate Accrued Interest and Market Value Sum for the current asset type
            Dim accruedInterestMarketValueSum As Decimal = dataTable.AsEnumerable().
        Where(Function(r) r.Field(Of String)("SRRe Asset Funding") = "Funded" AndAlso
                          r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                          r.Field(Of String)("SRRe Asset Type") = assetType AndAlso
                          r.Field(Of String)("SRRe IG") = "BIG" AndAlso
                          r.Field(Of String)("As Of Date") = asofdate).
        Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Accrued Interest"))) +
                       Convert.ToDecimal(Convert.ToDouble(r("Market Value"))))

            ' Write results to Excel
            wsBMA.Cell(rowIndexBMA, 1).Value = "Accrued Interest & Market Value FWH - " & assetType
            wsBMA.Cell(rowIndexBMA, 2).Value = accruedInterestMarketValueSum
            rowIndexBMA += 1
        Next
        'OC

        For Each assetType As String In srreAssetTypes
            ' Calculate Accrued Interest and Market Value Sum for the current asset type
            Dim accruedInterestMarketValueSum As Decimal = dataTable.AsEnumerable().
        Where(Function(r) r.Field(Of String)("SRRe Asset Funding") = "Funded" AndAlso
                          r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                          r.Field(Of String)("SRRe Asset Type") = assetType AndAlso
                           r.Field(Of String)("SRRe IG") = "IG" AndAlso
                          r.Field(Of String)("As Of Date") = asofdate).
        Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Accrued Interest"))) +
                       Convert.ToDecimal(Convert.ToDouble(r("Market Value"))))

            ' Write results to Excel
            wsBMA.Cell(rowIndexBMA, 1).Value = "Accrued Interest & Market Value OC - " & assetType
            wsBMA.Cell(rowIndexBMA, 2).Value = accruedInterestMarketValueSum
            rowIndexBMA += 1
        Next
        'BIG
        For Each assetType As String In srreAssetTypes
            ' Calculate Accrued Interest and Market Value Sum for the current asset type
            Dim accruedInterestMarketValueSum As Decimal = dataTable.AsEnumerable().
        Where(Function(r) r.Field(Of String)("SRRe Asset Funding") = "Funded" AndAlso
                          r.Field(Of String)("SRRe Port Code") = "OC" AndAlso
                          r.Field(Of String)("SRRe Asset Type") = assetType AndAlso
                           r.Field(Of String)("SRRe IG") = "BIG" AndAlso
                          r.Field(Of String)("As Of Date") = asofdate).
        Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Accrued Interest"))) +
                       Convert.ToDecimal(Convert.ToDouble(r("Market Value"))))

            ' Write results to Excel
            wsBMA.Cell(rowIndexBMA, 1).Value = "Accrued Interest & Market Value OC - " & assetType
            wsBMA.Cell(rowIndexBMA, 2).Value = accruedInterestMarketValueSum
            rowIndexBMA += 1
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
            newFileName = "D:\BMA_TP " & formattedDate & ".xlsx"

            Console.WriteLine(newFileName) ' Output: D:\BMA_TP11_30_2024.xlsx
        Else
            Console.WriteLine("Invalid date format.")
        End If

        ' Combine directory path with the new filename
        'Dim newFilePath As String = Path.Combine(directoryPath, newFileName)

        Console.WriteLine(newFileName)

        wbBMA.SaveAs(newFileName)
        Console.WriteLine("Report generated at: " & outputPath)
    End Sub
    Sub InsertPercentage(ByVal categoryName As String, ByVal categoryTotal As Integer, ByVal grandTotal As Integer)
        Dim percentage As Double = If(grandTotal > 0, (categoryTotal / grandTotal) * 100, 0)
        ws.Cell(rowIndex, 1).Value = categoryName & " (%)"
        ws.Cell(rowIndex, 2).Value = Math.Round(percentage, 2) & "%" ' Round to 2 decimal places
        rowIndex += 1
    End Sub
End Class
