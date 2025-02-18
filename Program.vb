Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports ADOX


Module Module1
    Sub Main()
        Dim excelPath As String = "C:\Users\manik\Downloads\Excel1.xlsx"
        Dim dbPath As String = "D:\EmployeeDB.accdb"

        ' Create Access DB if not exists
        If Not File.Exists(dbPath) Then
            CreateAccessDatabase(dbPath)
        End If

        Console.Write("Enter the table name: ")
        Dim tableName As String = Console.ReadLine().Trim()

        CreateTableFromExcel(dbPath, excelPath, tableName)

        ' Insert Data into Access
        InsertDataFromExcel(dbPath, excelPath, tableName)
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
    Sub CreateTableFromExcel(ByVal dbPath As String, ByVal excelFilePath As String, ByVal tableName As String)
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
                Dim createTableSQL As String = "CREATE TABLE [" & tableName & "] (ID AUTOINCREMENT PRIMARY KEY, "
                For Each col As DataColumn In dt.Columns
                    createTableSQL &= "[" & col.ColumnName & "] TEXT(255), "
                Next
                createTableSQL = createTableSQL.TrimEnd(","c, " "c) & ")"

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
    Sub InsertDataFromExcel(ByVal dbPath As String, ByVal excelFilePath As String, ByVal tableName As String)
        Try
            ' Connect to Excel File
            Dim excelConnStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & excelFilePath & ";Extended Properties='Excel 12.0 Xml;HDR=YES;'"
            Using excelConn As New OleDbConnection(excelConnStr)
                excelConn.Open()

                ' Get the first sheet name
                Dim sheetName As String = GetFirstSheetName(excelConn)

                ' Read all data excluding headers
                Dim sql As String = "SELECT * FROM [" & sheetName & "]"
                Dim dt As New DataTable()

                Using da As New OleDbDataAdapter(sql, excelConn)
                    da.Fill(dt)
                End Using

                ' Insert each row into Access Database
                Using accessConn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath)
                    accessConn.Open()

                    For Each row As DataRow In dt.Rows
                        Dim insertSQL As String = "INSERT INTO [" & tableName & "] ("
                        Dim valuesSQL As String = "VALUES ("

                        ' Dynamically build column names and values
                        For Each col As DataColumn In dt.Columns
                            insertSQL &= "[" & col.ColumnName & "], "
                            valuesSQL &= "'" & row(col.ColumnName).ToString().Replace("'", "''") & "', "
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
                End Using

                Console.WriteLine("Data inserted successfully into 'Employees' table.")

            End Using
        Catch ex As Exception
            Console.WriteLine("Error inserting data: " & ex.Message)
        End Try
    End Sub
End Module
