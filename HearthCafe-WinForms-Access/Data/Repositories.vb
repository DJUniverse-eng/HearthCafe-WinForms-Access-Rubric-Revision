Imports System.Data
Imports System.Data.OleDb
Imports System.Security.Cryptography
Imports System.Text

Namespace HearthCafe
    Public Module SecurityService
        Public Sub CreatePassword(password As String, ByRef hash As String, ByRef salt As String)
            Dim saltBytes(15) As Byte
            RandomNumberGenerator.Fill(saltBytes)
            salt = Convert.ToBase64String(saltBytes)
            hash = Convert.ToBase64String(Rfc2898DeriveBytes.Pbkdf2(password, saltBytes, 150000, HashAlgorithmName.SHA256, 32))
        End Sub
        Public Function VerifyPassword(password As String, salt As String, expectedHash As String) As Boolean
            Try
                Dim actual = Rfc2898DeriveBytes.Pbkdf2(password, Convert.FromBase64String(salt), 150000, HashAlgorithmName.SHA256, 32)
                Return CryptographicOperations.FixedTimeEquals(actual, Convert.FromBase64String(expectedHash))
            Catch
                Return False
            End Try
        End Function
    End Module

    Public Module UserRepository
        Public Function CountUsers() As Integer
            Using cn = Db.GetConnection(), cmd As New OleDbCommand("SELECT COUNT(*) FROM Users", cn)
                Return Convert.ToInt32(cmd.ExecuteScalar())
            End Using
        End Function
        Public Sub AddUser(username As String, displayName As String, role As String, password As String)
            Dim hash As String = "", salt As String = ""
            SecurityService.CreatePassword(password, hash, salt)
            Using cn = Db.GetConnection(), cmd As New OleDbCommand("INSERT INTO Users (Username,DisplayName,UserRole,PasswordHash,PasswordSalt,IsActive) VALUES (?,?,?,?,?,?)", cn)
                cmd.Parameters.AddWithValue("@p1", username.Trim()) : cmd.Parameters.AddWithValue("@p2", displayName.Trim())
                cmd.Parameters.AddWithValue("@p3", role) : cmd.Parameters.AddWithValue("@p4", hash) : cmd.Parameters.AddWithValue("@p5", salt) : cmd.Parameters.AddWithValue("@p6", True)
                cmd.ExecuteNonQuery()
            End Using
        End Sub
        Public Function Authenticate(username As String, password As String) As AppUser
            Using cn = Db.GetConnection(), cmd As New OleDbCommand("SELECT Id,Username,DisplayName,UserRole,PasswordHash,PasswordSalt,IsActive FROM Users WHERE Username=?", cn)
                cmd.Parameters.AddWithValue("@p1", username.Trim())
                Using r = cmd.ExecuteReader()
                    If Not r.Read() OrElse Not Convert.ToBoolean(r("IsActive")) Then Return Nothing
                    If Not SecurityService.VerifyPassword(password, Convert.ToString(r("PasswordSalt")), Convert.ToString(r("PasswordHash"))) Then Return Nothing
                    Return New AppUser With {.Id = Convert.ToInt32(r("Id")), .Username = Convert.ToString(r("Username")), .DisplayName = Convert.ToString(r("DisplayName")), .Role = Convert.ToString(r("UserRole")), .IsActive = True}
                End Using
            End Using
        End Function
        Public Function GetUsers() As DataTable
            Using cn = Db.GetConnection(), da As New OleDbDataAdapter("SELECT Id,Username,DisplayName,UserRole,IsActive FROM Users ORDER BY Username", cn)
                Dim t As New DataTable() : da.Fill(t) : Return t
            End Using
        End Function
        Public Sub SetUserActive(id As Integer, active As Boolean)
            Using cn = Db.GetConnection(), cmd As New OleDbCommand("UPDATE Users SET IsActive=? WHERE Id=?", cn)
                cmd.Parameters.AddWithValue("@p1", active) : cmd.Parameters.AddWithValue("@p2", id) : cmd.ExecuteNonQuery()
            End Using
        End Sub
        Public Sub ResetPassword(id As Integer, password As String)
            Dim hash As String = "", salt As String = "" : SecurityService.CreatePassword(password, hash, salt)
            Using cn = Db.GetConnection(), cmd As New OleDbCommand("UPDATE Users SET PasswordHash=?,PasswordSalt=? WHERE Id=?", cn)
                cmd.Parameters.AddWithValue("@p1", hash) : cmd.Parameters.AddWithValue("@p2", salt) : cmd.Parameters.AddWithValue("@p3", id) : cmd.ExecuteNonQuery()
            End Using
        End Sub
    End Module

    Public Module MenuRepository
        Public Function GetMenu(Optional archived As Boolean = False, Optional search As String = "") As DataTable
            Using cn = Db.GetConnection(), cmd As New OleDbCommand("SELECT Id,ItemName,Category,Price,Description,ImagePath,IsAvailable,IsArchived FROM MenuItems WHERE IsArchived=? AND (ItemName LIKE ? OR Category LIKE ?) ORDER BY Category,ItemName", cn)
                cmd.Parameters.AddWithValue("@p1", archived) : cmd.Parameters.AddWithValue("@p2", "%" & search & "%") : cmd.Parameters.AddWithValue("@p3", "%" & search & "%")
                Using da As New OleDbDataAdapter(cmd)
                    Dim t As New DataTable() : da.Fill(t) : Return t
                End Using
            End Using
        End Function
        Public Function GetAvailableMenu() As DataTable
            Using cn = Db.GetConnection(), da As New OleDbDataAdapter("SELECT Id,ItemName,Category,Price FROM MenuItems WHERE IsArchived=False AND IsAvailable=True ORDER BY Category,ItemName", cn)
                Dim t As New DataTable() : da.Fill(t) : Return t
            End Using
        End Function
        Public Sub SaveProduct(p As MenuProduct)
            If p.Name.Trim() = "" Then Throw New ArgumentException("Item name is required.")
            If p.Category.Trim() = "" Then Throw New ArgumentException("Category is required.")
            If p.Price <= 0D Then Throw New ArgumentException("Price must be greater than zero.")
            Using cn = Db.GetConnection()
                Using check As New OleDbCommand("SELECT COUNT(*) FROM MenuItems WHERE ItemName=? AND Category=? AND Id<>? AND IsArchived=False", cn)
                    check.Parameters.AddWithValue("@p1", p.Name.Trim()) : check.Parameters.AddWithValue("@p2", p.Category.Trim()) : check.Parameters.AddWithValue("@p3", p.Id)
                    If Convert.ToInt32(check.ExecuteScalar()) > 0 Then Throw New ArgumentException("An active item with this name already exists in the category.")
                End Using
                If p.Id = 0 Then
                    Using cmd As New OleDbCommand("INSERT INTO MenuItems (ItemName,Category,Price,Description,ImagePath,IsAvailable,IsArchived) VALUES (?,?,?,?,?,?,?)", cn)
                        AddProductParameters(cmd, p, False) : cmd.ExecuteNonQuery()
                    End Using
                Else
                    Using cmd As New OleDbCommand("UPDATE MenuItems SET ItemName=?,Category=?,Price=?,Description=?,ImagePath=?,IsAvailable=?,IsArchived=? WHERE Id=?", cn)
                        AddProductParameters(cmd, p, True) : cmd.Parameters.AddWithValue("@p8", p.Id) : cmd.ExecuteNonQuery()
                    End Using
                End If
            End Using
        End Sub
        Private Sub AddProductParameters(cmd As OleDbCommand, p As MenuProduct, includeArchive As Boolean)
            cmd.Parameters.AddWithValue("@p1", p.Name.Trim()) : cmd.Parameters.AddWithValue("@p2", p.Category.Trim()) : cmd.Parameters.AddWithValue("@p3", p.Price)
            cmd.Parameters.AddWithValue("@p4", p.Description) : cmd.Parameters.AddWithValue("@p5", p.ImagePath) : cmd.Parameters.AddWithValue("@p6", p.IsAvailable) : cmd.Parameters.AddWithValue("@p7", p.IsArchived)
        End Sub
        Public Sub ArchiveProduct(id As Integer, archived As Boolean)
            Using cn = Db.GetConnection(), cmd As New OleDbCommand("UPDATE MenuItems SET IsArchived=? WHERE Id=?", cn)
                cmd.Parameters.AddWithValue("@p1", archived) : cmd.Parameters.AddWithValue("@p2", id) : cmd.ExecuteNonQuery()
            End Using
        End Sub
    End Module

    Public Module OrderRepository
        Public Function GetOrders(Optional search As String = "", Optional fromDate As DateTime? = Nothing, Optional toDate As DateTime? = Nothing) As DataTable
            Dim sql = "SELECT Id,OrderNo,CustomerName,CreatedAt,OrderStatus,Subtotal,Discount,Total,PaymentMethod,AmountPaid,ChangeAmount,CreatedBy FROM Orders WHERE IsArchived=False AND (OrderNo LIKE ? OR CustomerName LIKE ?)"
            Using cn = Db.GetConnection(), cmd As New OleDbCommand()
                cmd.Connection = cn : cmd.CommandText = sql
                cmd.Parameters.AddWithValue("@p1", "%" & search & "%") : cmd.Parameters.AddWithValue("@p2", "%" & search & "%")
                If fromDate.HasValue Then cmd.CommandText &= " AND CreatedAt>=?" : cmd.Parameters.AddWithValue("@p3", fromDate.Value.Date)
                If toDate.HasValue Then cmd.CommandText &= " AND CreatedAt<?" : cmd.Parameters.AddWithValue("@p4", toDate.Value.Date.AddDays(1))
                cmd.CommandText &= " ORDER BY CreatedAt DESC"
                Using da As New OleDbDataAdapter(cmd)
                    Dim t As New DataTable() : da.Fill(t) : Return t
                End Using
            End Using
        End Function
        Public Function CreateOrder(customer As String, notes As String, discount As Decimal, paymentMethod As String, paid As Decimal, createdBy As String, cart As DataTable) As String
            If cart.Rows.Count = 0 Then Throw New ArgumentException("Add at least one item to the order.")
            If discount < 0D Then Throw New ArgumentException("Discount cannot be negative.")
            Dim subtotal As Decimal = 0D
            For Each row As DataRow In cart.Rows
                subtotal += Convert.ToDecimal(row("LineTotal"))
            Next
            If discount > subtotal Then Throw New ArgumentException("Discount cannot exceed subtotal.")
            Dim total = subtotal - discount
            If paid < total Then Throw New ArgumentException("Amount paid is less than the total.")
            Dim orderNo = "HC-" & DateTime.Now.ToString("yyyyMMddHHmmss") & "-" & Guid.NewGuid().ToString("N").Substring(0, 4).ToUpperInvariant()
            Using cn = Db.GetConnection(), tx = cn.BeginTransaction()
                Try
                    Using cmd As New OleDbCommand("INSERT INTO Orders (OrderNo,CustomerName,OrderNotes,CreatedAt,OrderStatus,Subtotal,Discount,Total,PaymentMethod,AmountPaid,ChangeAmount,CreatedBy,IsArchived) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", cn, tx)
                        cmd.Parameters.AddWithValue("@p1", orderNo) : cmd.Parameters.AddWithValue("@p2", customer.Trim()) : cmd.Parameters.AddWithValue("@p3", notes.Trim())
                        cmd.Parameters.AddWithValue("@p4", DateTime.Now) : cmd.Parameters.AddWithValue("@p5", "Pending") : cmd.Parameters.AddWithValue("@p6", subtotal)
                        cmd.Parameters.AddWithValue("@p7", discount) : cmd.Parameters.AddWithValue("@p8", total) : cmd.Parameters.AddWithValue("@p9", paymentMethod)
                        cmd.Parameters.AddWithValue("@p10", paid) : cmd.Parameters.AddWithValue("@p11", paid - total) : cmd.Parameters.AddWithValue("@p12", createdBy) : cmd.Parameters.AddWithValue("@p13", False)
                        cmd.ExecuteNonQuery()
                    End Using
                    Dim orderId As Integer
                    Using cmd As New OleDbCommand("SELECT @@IDENTITY", cn, tx) : orderId = Convert.ToInt32(cmd.ExecuteScalar()) : End Using
                    For Each row As DataRow In cart.Rows
                        Using cmd As New OleDbCommand("INSERT INTO OrderItems (OrderId,MenuItemId,ItemName,UnitPrice,Quantity,LineTotal) VALUES (?,?,?,?,?,?)", cn, tx)
                            cmd.Parameters.AddWithValue("@p1", orderId) : cmd.Parameters.AddWithValue("@p2", Convert.ToInt32(row("MenuItemId")))
                            cmd.Parameters.AddWithValue("@p3", Convert.ToString(row("ItemName"))) : cmd.Parameters.AddWithValue("@p4", Convert.ToDecimal(row("UnitPrice")))
                            cmd.Parameters.AddWithValue("@p5", Convert.ToInt32(row("Quantity"))) : cmd.Parameters.AddWithValue("@p6", Convert.ToDecimal(row("LineTotal"))) : cmd.ExecuteNonQuery()
                        End Using
                    Next
                    tx.Commit() : Return orderNo
                Catch
                    tx.Rollback() : Throw
                End Try
            End Using
        End Function
        Public Sub SetStatus(id As Integer, status As String)
            If Not {"Pending", "Preparing", "Completed", "Cancelled"}.Contains(status) Then Throw New ArgumentException("Invalid order status.")
            Using cn = Db.GetConnection(), cmd As New OleDbCommand("UPDATE Orders SET OrderStatus=?,CompletedAt=? WHERE Id=?", cn)
                cmd.Parameters.AddWithValue("@p1", status)
                If status = "Completed" OrElse status = "Cancelled" Then cmd.Parameters.AddWithValue("@p2", DateTime.Now) Else cmd.Parameters.AddWithValue("@p2", DBNull.Value)
                cmd.Parameters.AddWithValue("@p3", id) : cmd.ExecuteNonQuery()
            End Using
        End Sub
        Public Sub Archive(id As Integer)
            Using cn = Db.GetConnection(), cmd As New OleDbCommand("UPDATE Orders SET IsArchived=True WHERE Id=?", cn)
                cmd.Parameters.AddWithValue("@p1", id) : cmd.ExecuteNonQuery()
            End Using
        End Sub
        Public Function GetOrderItems(orderId As Integer) As DataTable
            Using cn = Db.GetConnection(), da As New OleDbDataAdapter("SELECT ItemName,UnitPrice,Quantity,LineTotal FROM OrderItems WHERE OrderId=" & orderId.ToString(Globalization.CultureInfo.InvariantCulture), cn)
                Dim t As New DataTable() : da.Fill(t) : Return t
            End Using
        End Function
        Public Function GetReport(fromDate As DateTime, toDate As DateTime) As DataTable
            Using cn = Db.GetConnection(), cmd As New OleDbCommand("SELECT OrderNo,CustomerName,CreatedAt,OrderStatus,Subtotal,Discount,Total,PaymentMethod,AmountPaid,ChangeAmount FROM Orders WHERE IsArchived=False AND CreatedAt>=? AND CreatedAt<? ORDER BY CreatedAt", cn)
                cmd.Parameters.AddWithValue("@p1", fromDate.Date) : cmd.Parameters.AddWithValue("@p2", toDate.Date.AddDays(1))
                Using da As New OleDbDataAdapter(cmd)
                    Dim t As New DataTable() : da.Fill(t) : Return t
                End Using
            End Using
        End Function
        Public Function GetDashboardTotals() As DataTable
            Using cn = Db.GetConnection(), da As New OleDbDataAdapter("SELECT COUNT(*) AS OrderCount, SUM(Total) AS SalesTotal FROM Orders WHERE IsArchived=False AND OrderStatus<>'Cancelled' AND CreatedAt>=Date()", cn)
                Dim t As New DataTable() : da.Fill(t) : Return t
            End Using
        End Function
    End Module
End Namespace
