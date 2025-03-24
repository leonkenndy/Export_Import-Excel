Imports System.Data
Imports System.Data.OleDb
Imports System.Diagnostics.Metrics
Imports System.Globalization
Imports System.IO
Imports System.IO.Pipelines
'Imports System.Net.WebRequestMethods
Imports System.Text.RegularExpressions
Imports System.Threading
Imports ADOX


Module Module1
    Sub Main()
        Dim excelPath As String = "C:\Users\manik\Downloads\ME Holdings File - 2024-11-30 - Skyridge.xlsx"
        Dim dbPath As String = "D:\ME Holdings.accdb"
        Dim excelReportPath As String = "D:\ME_Holdings.xlsx"

        ' Create Access DB if not exists
        If Not File.Exists(dbPath) Then
            CreateAccessDatabase(dbPath)
        End If

        Dim tableName As String = "ME_Holdings"
        Dim deletedate As String = Path.GetFileNameWithoutExtension(excelPath).Trim().Split(New String() {"ME Holdings File - "}, StringSplitOptions.None).Last()
        Dim match As Match = Regex.Match(deletedate, "\d{4}-\d{2}-\d{2}")
        Dim extractedDate As String = If(match.Success, match.Value, "Date not found")

        Dim parsedDate As DateTime = DateTime.ParseExact(extractedDate, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture)

        ' Format to MM/dd/yyyy
        Dim formattedDate As String = parsedDate.ToString("MM/dd/yyyy")

        CreateTableFromExcel(dbPath, excelPath, tableName, vbCrLf)

        ' Delete Data into Access
        DeleteDataFromAccessifDuplicate(dbPath, excelPath, tableName, formattedDate)
        ' Insert from in Excel
        InsertData(formattedDate, excelPath, tableName, dbPath)

        'Adding Columns to DataBase
        Dim columnsToAdd As New List(Of String) From {
                                                        "SRRe Asset Funding",
                                                        "SRRe Asset Category",
                                                        "SRRe Asset Type",
                                                        "SRRe Asset Subtype",
                                                        "SRRe Rating Letter",
                                                        "SRRe IG",
                                                        "SRRe Rating Number",
                                                        "SRRe Port Code"
                                                        }

        AddColumnsToAccessTable(dbPath, tableName, columnsToAdd)

        UpdateAddedColumns(dbPath, tableName, columnsToAdd, formattedDate)
        'pacwest
        'Update pacwest
        Dim primaryIDList1 As New List(Of String) From {"68295@AA7", "35116@AA3", "35092*AA3", "81800*AA0", "81786*AA8", "16325#AA8", "19418@AA2", "74744#AA0", "90282*AA8"}
        Dim primaryIDList2 As New List(Of String) From {"68295@AA7U", "35116@AA3U", "35092*AA3U", "81800*AA0U", "81786@AA8U", "16325#AA8U", "19418@AA2U", "74744#AA0U", "90282*AA8U"}
        Updatepacwest(dbPath, tableName, primaryIDList1, primaryIDList2)

        'Report generation to a blank file I'll make it optional 

        'Make report in Excel

        '        IG Corp Public 
        'D16 -> Sum of Stat Carry Value 4 Condition 
        'when 
        'SRRe Port Code = FWH
        'SRRe Asset Type = Corp Sec Public
        'SRRe IG = IG
        'As of Date = ME_Holding_File date

        'IG Corp Non Public
        'D17 -> Sum of Stat Carry Value 4 Condition 
        'when 
        'SRRe Port Code = FWH
        'SRRe Asset Type = Corp Sec Private
        'SRRe IG = IG
        'Date = ME_Holding_File Date
        'D15 = D16 + D17

        '        IG Corp Public 
        'D20 -> Sum of Stat Carry Value 4 Condition 
        'when 
        'SRRe Port Code = FWH
        'SRRe Asset Type = Corp Sec Public
        'SRRe IG = BIG
        'Date = ME_Holding_File Date

        'IG Corp Non Public
        'D21 -> Sum of Stat Carry Value 4 Condition 
        'when 
        'SRRe Port Code = FWH
        'SRRe Asset Type = Corp Sec Private
        'SRRe IG = BIG
        'Date = ME_Holding_File Date
        'D19 = D20 + D21


        '        Column E Is with port code as OC 
        'D21 -> Sum of Stat Carry Value 4 Condition 
        'when 
        'SRRe Port Code = OC
        'SRRe Asset Type = Corp Sec Private
        'SRRe IG = BIG
        'Date = ME_Holding_File Date
        'D19 = D20 + D21


        '        IG CLOs 

        ' -> Sum of Stat Carry Value 4 Condition 
        'when 
        'SRRe Port Code = FWH
        'SRRe Asset Type = CLO
        'SRRe IG = IG
        'Date = ME_Holding_File Date

        'Non IG CLOs 

        ' -> Sum of Stat Carry Value 4 Condition 
        'when 
        'SRRe Port Code = FWH
        'SRRe Asset Type = CLO
        'SRRe IG = BIG
        'Date = ME_Holding_File Date

        'CLOS -> Sum of Stat Carry Value 4 Condition 
        'when 
        'SRRe Port Code = FWH
        'SRRe Asset Type = CLO
        'Date = ME_Holding_File Date


        'Dont Check SRRe IG for 
        'Commercial paper 
        'Collateral loans 
        'syndicated bank loans 
        'ABS
        '        CMBS
        '        RMBS
        '        Direct Investments 
        'Preferred Stock 


        'Accrued Interest 
        'Add them based on 
        'ME_holding files date
        'FWH/ OC 

        Dim report As New ReportGenerator(dbPath, tableName, excelReportPath, formattedDate)
        report.GenerateReportFWH()
        'report.GenerateReportOC()
        Console.WriteLine("Data import completed successfully!")
        Console.ReadLine()
    End Sub

    ' Function to Create a New Access Database (Empty File)
    Sub CreateAccessDatabase(ByVal dbPath As String)
        Try
            ' Step 1: Create an empty Access database if it doesn't exist
            If Not File.Exists(dbPath) Then
                Dim cat As New Catalog()
                cat.Create("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";")
                Console.WriteLine("Database created successfully.")
            End If

        Catch ex As Exception
            Console.WriteLine("Error creating database: " & ex.Message)
        End Try
    End Sub


    ' Step 2: Read First Row of Excel & Create Table in Access
    Sub CreateTableFromExcel(ByVal dbPath As String, ByVal excelFilePath As String, ByVal tableName As String, vbCrLf As String)
        Try
            ' Connect to Excel File
            Dim excelConnStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & excelFilePath & ";Extended Properties='Excel 12.0 Xml;HDR=YES;'"
            Using excelConn As New OleDbConnection(excelConnStr)
                excelConn.Open()

                ' Read column names from the first row of the first sheet
                Dim sheetName As String = GetFirstSheetName(excelConn)
                Dim sql As String = "SELECT TOP 1 * FROM [" & sheetName & "]"
                Dim dt As New DataTable()

                Using da As New OleDbDataAdapter(sql, excelConn)
                    da.Fill(dt)
                End Using


                ' Generate CREATE TABLE statement dynamically
                Dim primaryKeys As String = "As Of Date ,Port Code "
                Dim createTableSQL As String = "CREATE TABLE [" & tableName & "] (" & vbCrLf

                ' Add all columns dynamically
                For Each col As DataColumn In dt.Columns
                    createTableSQL &= "[" & col.ColumnName & "] TEXT, " & vbCrLf
                Next

                ' Remove the last comma and newline before adding PRIMARY KEY
                createTableSQL = createTableSQL.TrimEnd(","c, " "c, vbCr, vbLf) & ")"

                ' Output the generated SQL (for debugging)
                Console.WriteLine("{0} (Length: {1})", createTableSQL, createTableSQL.Length)

                ' Execute CREATE TABLE in Access
                Using accessConn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath)
                    accessConn.Open()
                    Using cmd As New OleDbCommand(createTableSQL, accessConn)
                        cmd.ExecuteNonQuery()
                    End Using
                End Using

                Console.WriteLine("Table 'Employees' created successfully from Excel headers.")

            End Using
        Catch ex As OleDbException
            ' Logs specific OleDb SQL error
            Console.WriteLine("Database Error: " & ex.Message)
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.Message)
        End Try
    End Sub

    ' Step 3: Get the First Sheet Name
    Function GetFirstSheetName(excelConn As OleDbConnection) As String
        Dim dt As DataTable = excelConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
        If dt.Rows.Count > 0 Then
            Return dt.Rows(0)("TABLE_NAME").ToString()
        Else
            Throw New Exception("No sheets found in Excel file.")
        End If
    End Function
    ' Function to Import Excel Data into Access

    ' Step 4: Insert Data into Access Table from Excel
    Sub DeleteDataFromAccessifDuplicate(ByVal dbPath As String, ByVal excelFilePath As String, ByVal tableName As String, ByVal deleteDate As String)
        Try

            Using accessConn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath)
                accessConn.Open()

                Dim deleteSQL As String = "DELETE FROM [" & tableName & "] WHERE [As Of Date] = ? "
                Using deleteCmd As New OleDbCommand(deleteSQL, accessConn)
                    deleteCmd.Parameters.AddWithValue("?", deleteDate)
                    Dim count As Integer = deleteCmd.ExecuteNonQuery()
                    Console.WriteLine("Delete Number Of Record ", count)

                End Using
                accessConn.Close()

            End Using
            Console.WriteLine("Data updated successfully.")
        Catch ex As Exception
            Console.WriteLine("Error inserting data: " & ex.Message)
        End Try
    End Sub

    Sub InsertData(ByVal deleteDate As String, ByVal excelFilePath As String, ByVal tableName As String, ByVal dbPath As String)

        Dim excelConnStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & excelFilePath & ";Extended Properties='Excel 12.0 Xml;HDR=YES;'"
        Using excelConn As New OleDbConnection(excelConnStr)
            excelConn.Open()

            ' Read column names from the first row of the first sheet
            Dim sheetName As String = GetFirstSheetName(excelConn)
            Dim sql As String = "SELECT * FROM [" & sheetName & "] WHERE [Asset Category] IS NOT NULL AND [As Of Date] IS NOT NULL"
            Dim dt As New DataTable()

            Using da As New OleDbDataAdapter(sql, excelConn)
                da.Fill(dt)
            End Using

            ' Insert new or updated record

            Using accessConn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath)
                accessConn.Open()
                ' Dynamically build column names and values
                For Each row As DataRow In dt.Rows
                    Dim insertSQL As String = "INSERT INTO [" & tableName & "] ("
                    Dim valuesSQL As String = "VALUES ("
                    For Each col As DataColumn In dt.Columns
                        Dim colName As String = col.ColumnName.Trim() ' Trim column names to prevent issues
                        insertSQL &= "[" & colName & "], "
                        If (colName = "As Of Date") Then
                            valuesSQL &= "'" & deleteDate.ToString().Replace("'", "''") & "', "
                        Else
                            valuesSQL &= "'" & row(colName).ToString().Replace("'", "''") & "', "

                        End If
                    Next

                    insertSQL = insertSQL.TrimEnd(","c, " "c) & ")"
                    valuesSQL = valuesSQL.TrimEnd(","c, " "c) & ")"
                    ' Final SQL Statement
                    Dim finalSQL As String = insertSQL & " " & valuesSQL

                    ' Execute the insert command
                    Using cmd As New OleDbCommand(finalSQL, accessConn)
                        cmd.ExecuteNonQuery()
                    End Using
                Next




                Console.WriteLine("Data updated successfully.")
            End Using
        End Using
    End Sub

    Sub AddColumnsToAccessTable(ByVal dbPath As String, ByVal tableName As String, ByVal columnNames As List(Of String))
        ' Connection string
        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath

        Using conn As New OleDbConnection(connectionString)
            Try
                conn.Open()

                For Each columnName As String In columnNames
                    ' SQL command to add column
                    Dim sql As String = "ALTER TABLE [" & tableName & "] ADD COLUMN [" & columnName & "] TEXT(50);"

                    Using cmd As New OleDbCommand(sql, conn)
                        cmd.ExecuteNonQuery()
                        Console.WriteLine("Column '" & columnName & "' added successfully!")
                    End Using
                Next

            Catch ex As Exception
                Console.WriteLine("Error: " & ex.Message)
            End Try
        End Using
    End Sub


    Sub UpdateAddedColumns(ByVal dbPath As String, ByVal tableName As String, ByVal columnNames As List(Of String), ByVal specificdate As String)
        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath

        Using conn As New OleDbConnection(connectionString)
            Try
                conn.Open()
                For Each columnName As String In columnNames
                    Dim sql As String = ""
                    Dim sql2 As String = ""
                    Dim sql3 As String = ""

                    If columnName = "SRRe Asset Funding" Then
                        sql = "UPDATE [" & tableName & "] " &
                              "SET [" & columnName & "] = IIF([Asset Category] = 'Unfunded', 'Unfunded', 'Funded') " &
                              "WHERE [Asset Category] IS NOT NULL AND [As Of Date] IS NOT NULL;"
                    ElseIf columnName = "SRRe Asset Category" Then
                        'TODO infomation in excel seems incomplete
                        sql = "UPDATE [" & tableName & "] " &
                              "SET  [" & columnName & "] = " &
                              "IIf([Asset Type] IN ('IG Corp Public', 'Non IG Corp Public', 'Syndicated Bank Loans'), 'Corp Sec Public', " &
                              "IIf([Asset Type] IN ('IG Corp Non Public', 'Non-IG Corp Non Public', 'Commercial Mortgage Loans', 'Collateral Loans'), 'Private Credit', " &
                              "IIf([Asset Type] = 'Preferred Stock', 'Preferred Stock', " &
                              "IIf([Asset Type] = 'Direct Investments', 'Non-Fixed Income', " &
                              "IIf([Asset Type] IN ('CLO', 'ABS', 'CMBS', 'RMBS'), 'Structured Credit', [Asset Category])" &
                              ")))) " &
                              "WHERE [Asset Type] IS NOT NULL;"
                    ElseIf columnName = "SRRe Asset Type" Then
                        'TODO: its not Private Placement  its Private Placement Ind
                        sql = "UPDATE [" & tableName & "] " &
                          "SET [" & columnName & "] = " &
                          "IIF([Asset Type] = 'IG Corp Public' OR [Asset Type] = 'Non-IG Corp Public', 'Corp Sec Public', " &
                          "IIF(([Asset Type] = 'IG Corp Non Public' OR [Asset Type] = 'Non-IG Corp Non Public') AND [Private Placement Ind] = 'No', 'Corp Sec Private', " &
                          "IIF(([Asset Type] = 'IG Corp Non Public' OR [Asset Type] = 'Non-IG Corp Non Public') AND [Private Placement Ind] = 'Yes', 'Corp Sec 144a', " &
                          "[Asset Type])))" &
                          "WHERE [Asset Category] IS NOT NULL " &
                          "AND [As Of Date] IS NOT NULL;"
                    ElseIf columnName = "SRRe Asset Subtype" Then
                        sql = "UPDATE [" & tableName & "] " &
                          "SET [" & columnName & "] = " &
                          "IIF([Asset Subtype] = 'IG Corp Public' OR [Asset Subtype] = 'Non-IG Corp Public', 'Corp Sec Public', " &
                          "IIF(([Asset Subtype] = 'IG Corp Non Public' OR [Asset Subtype] = 'Non-IG Corp Non Public') AND [Private Placement Ind] = 'No', 'Corp Sec Private', " &
                          "IIF(([Asset Subtype] = 'IG Corp Non Public' OR [Asset Subtype] = 'Non-IG Corp Non Public') AND [Private Placement Ind] = 'Yes', 'Corp Sec 144a', " &
                          "[Asset Subtype])))" &
                          "WHERE [Asset Category] IS NOT NULL " &
                          "AND [As Of Date] IS NOT NULL;"
                        'Still Pending 
                    ElseIf columnName = "SRRe Rating Letter" Then
                        sql = "UPDATE [" & tableName & "] SET [" & columnName & "] = " &
                                "SWITCH(" &
                                "Trim([NRSRO Rating])='AAA', 'AAA', " &
                                "[NRSRO Rating] IN ('AA+', 'AA', 'AA-'), 'AA', " &
                                "[NRSRO Rating] IN ('A+', 'A', 'A-'), 'A', " &
                                "[NRSRO Rating] IN ('BBB+', 'BBB', 'BBB-'), 'BBB', " &
                                "[NRSRO Rating] IN ('BB+', 'BB', 'BB-'), 'BB', " &
                                "[NRSRO Rating] IN ('B+', 'B', 'B-'), 'B', " &
                                "[NRSRO Rating] IN ('CCC+', 'CCC', 'CCC-'), 'CCC', " &
                                "[NRSRO Rating] IN ('CC+', 'CC', 'CC-'), 'CC', " &
                                "[NRSRO Rating] IN ('C+', 'C', 'C-'), 'C', " &
                                "[NRSRO Rating]='D', 'D', " &
                                "True, NULL)"
                        sql2 = "UPDATE [" & tableName & "] SET [" & columnName & "] = " &
                                "SWITCH(" &
                                "Trim([BSCR Rating])='1','AAA'," &
                                "Trim([BSCR Rating])='2', 'AA', " &
                                "Trim([BSCR Rating])='3', 'A', " &
                                "Trim([BSCR Rating])='4', 'BBB', " &
                                "Trim([BSCR Rating])='5', 'BB', " &
                                "Trim([BSCR Rating])='6', 'B', " &
                                "Trim([BSCR Rating])='7', 'CCC', " &
                                "Trim([BSCR Rating])='8', 'NR', " &
                                "True, [" & columnName & "]) " &
                                "WHERE [" & columnName & "] IS NULL AND [BSCR Rating] IS NOT NULL ;"
                        sql3 = "UPDATE [" & tableName & "] " &
                                "SET [" & columnName & "] = " &
                                "SWITCH( " &
                                "Trim([CM])= 'CM1' , 'AAA', " &
                                "Trim([CM])=  'CM2' , 'BBB', " &
                                "Trim([CM])=  'CM3' , 'BB', " &
                                "Trim([CM])=  'CM4' , 'B', " &
                                "Trim([CM])=  'CM5' , 'CCC', " &
                                "Trim([CM])=  'CM6' , 'CC' " &
                                ") " &
                                "WHERE [" & columnName & "] IS NULL OR [BSCR Rating] IS NULL AND [CM] IS NOT NULL;"

                    ElseIf columnName = "SRRe IG" Then
                        sql = "UPDATE [" & tableName & "] " &
                               "SET [" & columnName & "] = " &
                               "SWITCH( " &
                               "Trim([SRRe Rating Letter])= 'AAA' , 'IG', " &
                               "Trim([SRRe Rating Letter])= 'AA' , 'IG', " &
                               "Trim([SRRe Rating Letter])= 'A' , 'IG', " &
                               "Trim([SRRe Rating Letter])=  'BBB' , 'IG', " &
                               "Trim([SRRe Rating Letter])=  'BB' , 'BIG', " &
                               "Trim([SRRe Rating Letter])=  'B' , 'BIG', " &
                               "Trim([SRRe Rating Letter])=  'CCC' , 'BIG', " &
                               "Trim([SRRe Rating Letter])= 'CC' , 'BIG', " &
                               "Trim([SRRe Rating Letter])=  'C' , 'BIG' " &
                               ") " &
                               "WHERE [" & columnName & "] IS NULL OR [SRRe Rating Letter] IS NOT NULL;"
                    ElseIf columnName = "SRRe Rating Number" Then
                        sql = "UPDATE [" & tableName & "] " &
                               "SET [" & columnName & "] = " &
                               "SWITCH( " &
                               "Trim([SRRe Rating Letter])= 'AAA' , '1', " &
                               "Trim([SRRe Rating Letter])= 'AA' , '2', " &
                               "Trim([SRRe Rating Letter])= 'A' , '3', " &
                               "Trim([SRRe Rating Letter])=  'BBB' , '4', " &
                               "Trim([SRRe Rating Letter])=  'BB' , '5', " &
                               "Trim([SRRe Rating Letter])=  'B' , '6', " &
                               "Trim([SRRe Rating Letter])=  'CCC' , '7', " &
                                "Trim([SRRe Rating Letter])= 'CC' , '8', " &
                               "Trim([SRRe Rating Letter])=  'C' , '9' " &
                               ") " &
                               "WHERE [" & columnName & "] IS NULL OR [SRRe Rating Letter] IS NOT NULL;"
                    ElseIf columnName = "SRRe Port Code" Then
                        sql = "UPDATE [" & tableName & "] " &
                               "SET [" & columnName & "] = " &
                               "SWITCH( " &
                               "Trim([Port Code])= '811' , 'FWH', " &
                               "Trim([Port Code])= '812' , 'OC', " &
                               "Trim([Port Code])= '813' , 'FWH', " &
                               "Trim([Port Code])= '814' , 'FWH', " &
                               "Trim([Port Code])= '815' , 'FWH', " &
                               "Trim([Port Code])= '816' , 'FWH', " &
                               "Trim([Port Code])= '817' , 'OC', " &
                               "Trim([Port Code])= '821' , 'OC', " &
                               "Trim([Port Code])= '823' , 'OC' " &
                               ") " &
                               "WHERE [" & columnName & "] IS NULL OR [Port Code] IS NOT NULL;"
                    End If


                    ' Execute only if SQL is not empty
                    If sql <> "" Then
                        Using cmd As New OleDbCommand(sql, conn)
                            Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                            Console.WriteLine("Executing 1st SQL: " & vbCrLf & sql)
                            Console.WriteLine($"{columnName}: {rowsAffected} rows updated successfully!")
                        End Using
                    End If
                    If sql2 <> "" Then
                        Using cmd As New OleDbCommand(sql2, conn)
                            Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                            Console.WriteLine("Executing 2nd SQL: " & vbCrLf & sql2)
                            Console.WriteLine($"{columnName}: {rowsAffected} rows updated successfully!")
                        End Using
                    End If
                    If sql3 <> "" Then
                        Using cmd As New OleDbCommand(sql3, conn)
                            Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                            Console.WriteLine("Executing 3rd SQL: " & vbCrLf & sql3)
                            Console.WriteLine($"{columnName}: {rowsAffected} rows updated successfully!")
                        End Using
                    End If
                Next
            Catch ex As OleDbException
                ' Log General Error Message
                Console.WriteLine("OleDb Database Error: " & ex.Message)
            Catch ex As Exception
                ' General Error Logging
                Console.WriteLine("General Error: " & ex.Message)

            End Try
        End Using
    End Sub

    Public Sub Updatepacwest(ByVal dbPath As String, ByVal tableName As String, primaryID1 As List(Of String), primaryID2 As List(Of String))
        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath

        Using conn As New OleDbConnection(connectionString)
            Try
                conn.Open()

                Dim combinedSQLs As New List(Of String)()

                ' ---- First Primary ID List - Funded ----
                If primaryID1 IsNot Nothing AndAlso primaryID1.Count > 0 Then
                    Dim formattedIDs As String = String.Join(", ", primaryID1.Select(Function(id) "'" & id & "'"))
                    Dim whereClause As String = "WHERE [Primary ID] IN (" & formattedIDs & ")"
                    Dim sql1 As String = "UPDATE [" & tableName & "] " &
                        "SET [SRRe Asset Funding] = 'Funded', " &
                        "[SRRe Asset Category] = 'Private Credit', " &
                        "[SRRe Asset Type] = 'Construction Loan', " &
                        "[SRRe Asset Subtype] = 'Construction Loan', " &
                        "[SRRE Rating Letter] = 'BBB', " &
                        "[SRRE Rating Number] = '4' " & whereClause & ";"
                    combinedSQLs.Add(sql1)
                End If

                ' ---- Second Primary ID List - Unfunded ----
                If primaryID2 IsNot Nothing AndAlso primaryID2.Count > 0 Then
                    Dim formattedIDs As String = String.Join(", ", primaryID2.Select(Function(id) "'" & id & "'"))
                    Dim whereClause As String = "WHERE [Primary ID] IN (" & formattedIDs & ")"
                    Dim sql2 As String = "UPDATE [" & tableName & "] " &
                        "SET [SRRe Asset Funding] = 'Unfunded', " &
                        "[SRRe Asset Category] = 'Private Credit', " &
                        "[SRRe Asset Type] = 'Construction Loan', " &
                        "[SRRe Asset Subtype] = 'Construction Loan', " &
                        "[SRRE Rating Letter] = 'BBB', " &
                        "[SRRE Rating Number] = '4' " & whereClause & ";"
                    combinedSQLs.Add(sql2)
                End If

                ' ---- Execute All Queries ----
                For Each sql As String In combinedSQLs
                    Using cmd As New OleDbCommand(sql, conn)
                        cmd.ExecuteNonQuery()
                    End Using
                Next

            Catch ex As OleDbException
                Console.WriteLine("OleDb Database Error: " & ex.Message)
            Catch ex As Exception
                Console.WriteLine("General Error: " & ex.Message)
            End Try
        End Using
    End Sub


End Module
