Imports System.Data
Imports System.Data.OleDb
Imports System.Globalization
Imports System.IO
Imports System.Text.RegularExpressions
Imports ADOX


Module Module1
    Sub Main()
        Dim excelPath As String = "C:\Users\manik\Downloads\ME Holdings File - 2024-11-30 - Skyridge.xlsx"
        Dim dbPath As String = "D:\ME Holdings.accdb"

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
        InsertData(deletedate, excelPath, tableName, dbPath)

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

                Dim deleteSQL As String = "DELETE FROM [" & tableName & "] WHERE [As Of Date] = ?"
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
            Dim sql As String = "SELECT * FROM [" & sheetName & "]"
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
End Module
