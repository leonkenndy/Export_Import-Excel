Imports System.Data.OleDb
Imports ClosedXML.Excel

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

    Public Sub GenerateReportFWH()
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

        ws.Cell(rowIndex, 1).Value = "Cash and Equivalents - Pre-QSSA"
        ws.Cell(rowIndex, 2).Value = cashEquivalentPreQssa
        rowIndex += 1

        'Cash and Equivalents
        Dim cashEquivalent = dataTable.AsEnumerable().Where(Function(r) r.Field(Of String)("SRRe Port Code") = "FWH" AndAlso
                                                             r.Field(Of String)("SRRe Asset Type") = "Commercial Paper").Sum(Function(r) Convert.ToDecimal(Convert.ToDouble(r("Stat Carry Value"))))

        ws.Cell(rowIndex, 1).Value = "Cash and Equivalents"
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

        Dim TOTAL_Cash_and_Equivalents As Integer = cashEquivalent + cashEquivalentOC
        Dim TOTAL_Cash_Equivalent_ As Integer = cashEquivalentPreQssa + cashEquivalentPreQssaOC
        Dim TOTAL_commercialPaper As Integer = commercialPaper + commercialPaperOC
        Dim TOTAL_directInvestments As Integer = directInvestments + directInvestmentsOC
        Dim TOTAL_igCorpPublicSum As Integer = igCorpPublicSum + igCorpPublicSumOC

        Dim TOTAL_igCorpSumOC As Integer = igCorpSumOC + igCorpSum
        Dim TOTAL_igCorpNonPublicSum As Integer = igCorpNonPublicSum + igCorpNonPublicSumOC

        Dim TOTAL_bigCorpSum As Integer = bigCorpSum + bigCorpSumOC
        Dim TOTAL_nonigcml As Integer = nonigcml + nonigcmlOC

        Dim TOTAL_nonIGCorpNonPublicSum As Integer = nonIGCorpNonPublicSum + nonIGCorpNonPublicSumOC
        Dim TOTAL_nonIGCorpPublicSum As Integer = nonIGCorpPublicSum + nonIGCorpPublicSumOC
        Dim TOTAL_NonIGCLOs As Integer = NonIGCLOs + NonIGCLOsOC
        Dim TOTAL_collateral_loans As Integer = collateral_loans + collateral_loansOC
        Dim TOTAL_syndicated_bank_loans As Integer = syndicated_bank_loans + syndicated_bank_loansOC
        Dim TOTAL_igcmls As Integer = igcmls + igcmlsOC
        Dim TOTAL_totalcmls As Integer = totalcmls + totalcmlsOC
        Dim TOTAL_othersStructureAssets As Integer = othersStructureAssets + othersStructureAssetsOC
        Dim TOTAL_others As Integer = others + othersOC
        Dim TOTAL_preferredStock As Integer = preferredStock + preferredStockOC
        '------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ' Accrued Interest 

        wb.SaveAs(outputPath)
        Console.WriteLine("Report generated at: " & outputPath)
    End Sub
End Class
