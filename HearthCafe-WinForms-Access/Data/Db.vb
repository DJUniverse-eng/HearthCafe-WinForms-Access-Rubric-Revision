Imports System.Data
Imports System.Data.OleDb
Imports System.IO

Namespace HearthCafe
    Public Module Db
        ' Change this path if your school requires the .accdb file in another location.
        Public ReadOnly DatabasePath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "HearthCafe.accdb")
        Private ReadOnly ConnectionString As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={DatabasePath};Persist Security Info=False;"

        Public Function GetConnection() As OleDbConnection
            Dim cn As New OleDbConnection(ConnectionString)
            cn.Open()
            Return cn
        End Function

        Public Sub Initialize()
            Dim folder = Path.GetDirectoryName(DatabasePath)
            If Not Directory.Exists(folder) Then Directory.CreateDirectory(folder)
            If Not File.Exists(DatabasePath) Then
                Try
                    Dim catalogType = Type.GetTypeFromProgID("ADOX.Catalog")
                    If catalogType Is Nothing Then Throw New InvalidOperationException("ADOX is not installed. Create a blank Access .accdb database named HearthCafe.accdb in the Data folder, then restart.")
                    Dim catalog = Activator.CreateInstance(catalogType)
                    catalogType.InvokeMember("Create", Reflection.BindingFlags.InvokeMethod, Nothing, catalog,
                        New Object() {ConnectionString})
                    Runtime.InteropServices.Marshal.FinalReleaseComObject(catalog)
                Catch ex As Exception
                    Throw New InvalidOperationException("The Access database could not be created automatically. Install Microsoft Access Database Engine (ACE), or create a blank Access database at: " & DatabasePath, ex)
                End Try
            End If

            CreateTableIfMissing("Users", "CREATE TABLE Users (Id COUNTER CONSTRAINT PK_Users PRIMARY KEY, Username TEXT(80) NOT NULL, DisplayName TEXT(120) NOT NULL, UserRole TEXT(20) NOT NULL, PasswordHash TEXT(128) NOT NULL, PasswordSalt TEXT(128) NOT NULL, IsActive YESNO NOT NULL)")
            CreateTableIfMissing("MenuItems", "CREATE TABLE MenuItems (Id COUNTER CONSTRAINT PK_MenuItems PRIMARY KEY, ItemName TEXT(120) NOT NULL, Category TEXT(80) NOT NULL, Price CURRENCY NOT NULL, Description MEMO, ImagePath TEXT(255), IsAvailable YESNO NOT NULL, IsArchived YESNO NOT NULL)")
            CreateTableIfMissing("Orders", "CREATE TABLE Orders (Id COUNTER CONSTRAINT PK_Orders PRIMARY KEY, OrderNo TEXT(30) NOT NULL, CustomerName TEXT(120), OrderNotes MEMO, CreatedAt DATETIME NOT NULL, CompletedAt DATETIME, OrderStatus TEXT(20) NOT NULL, Subtotal CURRENCY NOT NULL, Discount CURRENCY NOT NULL, Total CURRENCY NOT NULL, PaymentMethod TEXT(30), AmountPaid CURRENCY NOT NULL, ChangeAmount CURRENCY NOT NULL, CreatedBy TEXT(80), IsArchived YESNO NOT NULL)")
            CreateTableIfMissing("OrderItems", "CREATE TABLE OrderItems (Id COUNTER CONSTRAINT PK_OrderItems PRIMARY KEY, OrderId LONG NOT NULL, MenuItemId LONG, ItemName TEXT(120) NOT NULL, UnitPrice CURRENCY NOT NULL, Quantity LONG NOT NULL, LineTotal CURRENCY NOT NULL)")
            CreateTableIfMissing("AppSettings", "CREATE TABLE AppSettings (SettingKey TEXT(80) CONSTRAINT PK_AppSettings PRIMARY KEY, SettingValue MEMO)")
            EnsureIndex("Users", "UX_Users_Username", "CREATE UNIQUE INDEX UX_Users_Username ON Users (Username)")
            EnsureIndex("Orders", "UX_Orders_OrderNo", "CREATE UNIQUE INDEX UX_Orders_OrderNo ON Orders (OrderNo)")
            EnsureSetting("SystemName", "Hearth Café Ordering and Management System")
            EnsureSetting("Currency", "₱")
            EnsureSetting("ReceiptFooter", "Thank you for visiting Hearth Café!")
        End Sub

        Private Sub CreateTableIfMissing(tableName As String, sql As String)
            Using cn = GetConnection()
                Dim schema = cn.GetSchema("Tables")
                For Each row As DataRow In schema.Rows
                    If String.Equals(Convert.ToString(row("TABLE_NAME")), tableName, StringComparison.OrdinalIgnoreCase) Then Return
                Next
                Using cmd As New OleDbCommand(sql, cn) : cmd.ExecuteNonQuery() : End Using
            End Using
        End Sub

        Private Sub EnsureIndex(tableName As String, indexName As String, sql As String)
            Try
                Using cn = GetConnection()
                    Dim schema = cn.GetSchema("Indexes", New String() {Nothing, Nothing, Nothing, indexName})
                    If schema.Rows.Count = 0 Then Using cmd As New OleDbCommand(sql, cn) : cmd.ExecuteNonQuery() : End Using
                End Using
            Catch
                ' Unique constraints are still validated in the application if an index already exists under another name.
            End Try
        End Sub

        Public Sub EnsureSetting(key As String, value As String)
            Using cn = GetConnection()
                Using cmd As New OleDbCommand("SELECT COUNT(*) FROM AppSettings WHERE SettingKey=?", cn)
                    cmd.Parameters.AddWithValue("@p1", key)
                    If Convert.ToInt32(cmd.ExecuteScalar()) = 0 Then
                        Using ins As New OleDbCommand("INSERT INTO AppSettings (SettingKey, SettingValue) VALUES (?,?)", cn)
                            ins.Parameters.AddWithValue("@p1", key) : ins.Parameters.AddWithValue("@p2", value) : ins.ExecuteNonQuery()
                        End Using
                    End If
                End Using
            End Using
        End Sub

        Public Function GetSetting(key As String, Optional fallback As String = "") As String
            Using cn = GetConnection()
                Using cmd As New OleDbCommand("SELECT SettingValue FROM AppSettings WHERE SettingKey=?", cn)
                    cmd.Parameters.AddWithValue("@p1", key)
                    Dim result = cmd.ExecuteScalar()
                    Return If(result Is Nothing OrElse result Is DBNull.Value, fallback, Convert.ToString(result))
                End Using
            End Using
        End Function

        Public Sub SaveSetting(key As String, value As String)
            Using cn = GetConnection()
                Using cmd As New OleDbCommand("UPDATE AppSettings SET SettingValue=? WHERE SettingKey=?", cn)
                    cmd.Parameters.AddWithValue("@p1", value) : cmd.Parameters.AddWithValue("@p2", key)
                    If cmd.ExecuteNonQuery() = 0 Then
                        Using ins As New OleDbCommand("INSERT INTO AppSettings (SettingKey, SettingValue) VALUES (?,?)", cn)
                            ins.Parameters.AddWithValue("@p1", key) : ins.Parameters.AddWithValue("@p2", value) : ins.ExecuteNonQuery()
                        End Using
                    End If
                End Using
            End Using
        End Sub
    End Module
End Namespace
