Option Strict On
Option Explicit On
Option Infer On

Imports System
Imports System.Data
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Data.OleDb
Imports System.Windows.Forms

Namespace HearthCafe

    Public Class MainForm
        Inherits Form

        Private ReadOnly currentUser As AppUser
        Private ReadOnly tabs As New TabControl With {
            .Dock = DockStyle.Fill
        }

        Private menuGrid As DataGridView
        Private orderGrid As DataGridView
        Private reportGrid As DataGridView
        Private usersGrid As DataGridView

        Private menuSearch As TextBox
        Private orderSearch As TextBox
        Private archivedMenu As CheckBox

        Private menuCombo As ComboBox
        Private quantity As NumericUpDown
        Private cartGrid As DataGridView
        Private cart As DataTable

        Private customerName As TextBox
        Private orderNotes As TextBox
        Private discount As NumericUpDown
        Private amountPaid As NumericUpDown
        Private paymentMethod As ComboBox
        Private lblSubtotal As Label
        Private lblTotal As Label

        Private fromDate As DateTimePicker
        Private toDate As DateTimePicker
        Private reportSummary As Label

        Private settingName As TextBox
        Private settingFooter As TextBox
        Private settingCurrency As TextBox

        Public Sub New(user As AppUser)
            If user Is Nothing Then
                Throw New ArgumentNullException(NameOf(user))
            End If

            currentUser = user

            Text = "Hearth Café Ordering and Management System"
            WindowState = FormWindowState.Maximized
            MinimumSize = New Size(1100, 720)
            BackColor = Color.FromArgb(249, 244, 236)
            Font = New Font("Segoe UI", 9.5F)

            Dim header As New Panel With {
                .Dock = DockStyle.Top,
                .Height = 76,
                .BackColor = Color.FromArgb(95, 54, 35)
            }

            Dim title As New Label With {
                .Text = Db.GetSetting("SystemName", "HEARTH CAFÉ"),
                .ForeColor = Color.White,
                .Font = New Font("Segoe UI", 18, FontStyle.Bold),
                .AutoSize = True,
                .Location = New Point(22, 10)
            }

            Dim who As New Label With {
                .Text = user.DisplayName & "  |  " & user.Role,
                .ForeColor = Color.White,
                .AutoSize = True,
                .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
                .Location = New Point(780, 28)
            }

            Dim logout As New Button With {
                .Text = "Log Out",
                .Width = 95,
                .Height = 34,
                .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
                .Location = New Point(1010, 20)
            }

            AddHandler logout.Click,
                Sub(sender As Object, e As EventArgs)
                    If MessageBox.Show(
                        "Log out of Hearth Café?",
                        "Log Out",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question
                    ) = DialogResult.Yes Then
                        Close()
                    End If
                End Sub

            header.Controls.AddRange(New Control() {title, who, logout})

            Controls.Add(tabs)
            Controls.Add(header)

            AddTab("Dashboard", BuildDashboard())
            AddTab("Menu Management", BuildMenu())
            AddTab("New Order", BuildOrderEntry())
            AddTab("Order Management", BuildOrders())
            AddTab("Reports", BuildReports())

            If IsAdmin() Then
                AddTab("User Accounts", BuildUsers())
                AddTab("Settings", BuildSettings())
            End If

            AddHandler tabs.SelectedIndexChanged,
                Sub(sender As Object, e As EventArgs)
                    RefreshCurrentTab()
                End Sub

            RefreshAll()
        End Sub

        Private Sub AddTab(title As String, page As Control)
            Dim tab As New TabPage(title) With {
                .BackColor = Color.FromArgb(249, 244, 236),
                .Padding = New Padding(12)
            }

            page.Dock = DockStyle.Fill
            tab.Controls.Add(page)
            tabs.TabPages.Add(tab)
        End Sub

        '=========================================================
        ' DASHBOARD
        '=========================================================

        Private Function BuildDashboard() As Control
            Dim root As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 3,
                .Padding = New Padding(16)
            }

            root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
            root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))

            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 60))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 150))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100))

            Dim heading As New Label With {
                .Text = "DASHBOARD OVERVIEW",
                .Font = New Font("Segoe UI", 18, FontStyle.Bold),
                .AutoSize = True
            }

            root.Controls.Add(heading, 0, 0)
            root.SetColumnSpan(heading, 2)

            Dim ordersCard As New Label With {
                .Dock = DockStyle.Fill,
                .BackColor = Color.White,
                .Font = New Font("Segoe UI", 15, FontStyle.Bold),
                .TextAlign = ContentAlignment.MiddleCenter,
                .Text = "Today's Orders"
            }

            Dim salesCard As New Label With {
                .Dock = DockStyle.Fill,
                .BackColor = Color.White,
                .Font = New Font("Segoe UI", 15, FontStyle.Bold),
                .TextAlign = ContentAlignment.MiddleCenter,
                .Text = "Today's Sales"
            }

            root.Controls.Add(ordersCard, 0, 1)
            root.Controls.Add(salesCard, 1, 1)

            Dim recent As New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                .BackgroundColor = Color.White,
                .AllowUserToAddRows = False,
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                .MultiSelect = False
            }

            root.Controls.Add(recent, 0, 2)
            root.SetColumnSpan(recent, 2)

            AddHandler root.HandleCreated,
                Sub(sender As Object, e As EventArgs)
                    Try
                        Using cn As OleDbConnection = Db.GetConnection()
                            Dim sql As String =
                                "SELECT COUNT(*) AS OrderCount, " &
                                "SUM(Total) AS SalesTotal " &
                                "FROM Orders " &
                                "WHERE IsArchived=False " &
                                "AND OrderStatus<>'Cancelled' " &
                                "AND CreatedAt>=Date()"

                            Using da As New OleDbDataAdapter(sql, cn)
                                Dim t As New DataTable()
                                da.Fill(t)

                                Dim orderCount As Integer = 0
                                Dim salesTotal As Decimal = 0D

                                If t.Rows.Count > 0 Then
                                    If Not IsDBNull(t.Rows(0)("OrderCount")) Then
                                        orderCount =
                                            Convert.ToInt32(t.Rows(0)("OrderCount"))
                                    End If

                                    If Not IsDBNull(t.Rows(0)("SalesTotal")) Then
                                        salesTotal =
                                            Convert.ToDecimal(t.Rows(0)("SalesTotal"))
                                    End If
                                End If

                                ordersCard.Text =
                                    "TODAY'S ORDERS" &
                                    Environment.NewLine &
                                    orderCount.ToString()

                                salesCard.Text =
                                    "TODAY'S SALES" &
                                    Environment.NewLine &
                                    salesTotal.ToString(
                                        "C2",
                                        CultureInfo.GetCultureInfo("en-PH")
                                    )
                            End Using
                        End Using

                        Using cn As OleDbConnection = Db.GetConnection()
                            Dim sql As String =
                                "SELECT TOP 15 OrderNo, CustomerName, CreatedAt, " &
                                "OrderStatus, Total " &
                                "FROM Orders " &
                                "WHERE IsArchived=False " &
                                "ORDER BY CreatedAt DESC"

                            Using da As New OleDbDataAdapter(sql, cn)
                                Dim t As New DataTable()
                                da.Fill(t)
                                recent.DataSource = t
                            End Using
                        End Using

                    Catch ex As Exception
                        MessageBox.Show(
                            "Dashboard could not load: " & ex.Message,
                            "Dashboard Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning
                        )
                    End Try
                End Sub

            Return root
        End Function

        '=========================================================
        ' MENU MANAGEMENT
        '=========================================================

        Private Function BuildMenu() As Control
            Dim root As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .RowCount = 2,
                .ColumnCount = 1,
                .Padding = New Padding(8)
            }

            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 52))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100))

            Dim bar As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .WrapContents = False,
                .AutoScroll = True
            }

            menuSearch = New TextBox With {
                .Width = 240,
                .PlaceholderText = "Search name or category..."
            }

            Dim searchBtn As New Button With {
                .Text = "Search",
                .Width = 80
            }

            Dim add As New Button With {
                .Text = "Add Item",
                .Width = 90
            }

            Dim edit As New Button With {
                .Text = "Edit Selected",
                .Width = 110
            }

            Dim archive As New Button With {
                .Text = "Archive / Restore",
                .Width = 130
            }

            archivedMenu = New CheckBox With {
                .Text = "Show archived",
                .AutoSize = True,
                .Padding = New Padding(4, 8, 4, 0)
            }

            bar.Controls.AddRange(
                New Control() {
                    menuSearch, searchBtn, add, edit, archive, archivedMenu
                }
            )

            menuGrid = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                .MultiSelect = False,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                .AllowUserToAddRows = False,
                .BackgroundColor = Color.White
            }

            root.Controls.Add(bar, 0, 0)
            root.Controls.Add(menuGrid, 0, 1)

            AddHandler searchBtn.Click,
                Sub(sender As Object, e As EventArgs)
                    LoadMenuGrid()
                End Sub

            AddHandler menuSearch.KeyDown,
                Sub(sender As Object, e As KeyEventArgs)
                    If e.KeyCode = Keys.Enter Then
                        LoadMenuGrid()
                        e.SuppressKeyPress = True
                    End If
                End Sub

            AddHandler archivedMenu.CheckedChanged,
                Sub(sender As Object, e As EventArgs)
                    LoadMenuGrid()
                End Sub

            AddHandler add.Click,
                Sub(sender As Object, e As EventArgs)
                    If Not IsAdmin() Then
                        Denied()
                        Return
                    End If

                    Using f As New ProductEditorForm()
                        If f.ShowDialog(Me) = DialogResult.OK Then
                            LoadMenuGrid()
                            ReloadOrderMenu()
                        End If
                    End Using
                End Sub

            AddHandler edit.Click, AddressOf EditProduct

            AddHandler menuGrid.CellDoubleClick,
                Sub(sender As Object, e As DataGridViewCellEventArgs)
                    If e.RowIndex >= 0 Then
                        EditProduct(sender, EventArgs.Empty)
                    End If
                End Sub

            AddHandler archive.Click, AddressOf ToggleArchiveProduct

            Return root
        End Function

        Private Sub LoadMenuGrid()
            If menuGrid Is Nothing Then Return

            Try
                menuGrid.DataSource =
                    MenuRepository.GetMenu(
                        archivedMenu.Checked,
                        menuSearch.Text.Trim()
                    )

                If menuGrid.Columns.Contains("Id") Then
                    menuGrid.Columns("Id").Visible = False
                End If

                If menuGrid.Columns.Contains("ImagePath") Then
                    menuGrid.Columns("ImagePath").Visible = False
                End If

                If menuGrid.Columns.Contains("Description") Then
                    menuGrid.Columns("Description").Visible = False
                End If

                If menuGrid.Columns.Contains("Price") Then
                    menuGrid.Columns("Price").DefaultCellStyle.Format = "C2"
                End If

            Catch ex As Exception
                MessageBox.Show(
                    "Menu could not load: " & ex.Message,
                    "Menu Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                )
            End Try
        End Sub

        Private Function SelectedMenuProduct() As MenuProduct
            If menuGrid Is Nothing OrElse menuGrid.CurrentRow Is Nothing Then
                Return Nothing
            End If

            Dim r As DataGridViewRow = menuGrid.CurrentRow

            If Not menuGrid.Columns.Contains("Id") Then Return Nothing

            If r.Cells("Id").Value Is Nothing OrElse
               IsDBNull(r.Cells("Id").Value) Then
                Return Nothing
            End If

            Return New MenuProduct With {
                .Id = Convert.ToInt32(r.Cells("Id").Value),
                .Name = Convert.ToString(r.Cells("ItemName").Value),
                .Category = Convert.ToString(r.Cells("Category").Value),
                .Price = Convert.ToDecimal(r.Cells("Price").Value),
                .Description = Convert.ToString(r.Cells("Description").Value),
                .ImagePath = Convert.ToString(r.Cells("ImagePath").Value),
                .IsAvailable = Convert.ToBoolean(r.Cells("IsAvailable").Value),
                .IsArchived = Convert.ToBoolean(r.Cells("IsArchived").Value)
            }
        End Function

        Private Sub EditProduct(sender As Object, e As EventArgs)
            If Not IsAdmin() Then
                Denied()
                Return
            End If

            Dim p As MenuProduct = SelectedMenuProduct()

            If p Is Nothing Then
                MessageBox.Show("Select a menu item first.")
                Return
            End If

            Using f As New ProductEditorForm(p)
                If f.ShowDialog(Me) = DialogResult.OK Then
                    LoadMenuGrid()
                    ReloadOrderMenu()
                End If
            End Using
        End Sub

        Private Sub ToggleArchiveProduct(sender As Object, e As EventArgs)
            If Not IsAdmin() Then
                Denied()
                Return
            End If

            Dim p As MenuProduct = SelectedMenuProduct()

            If p Is Nothing Then
                MessageBox.Show("Select a menu item first.")
                Return
            End If

            Dim question As String

            If p.IsArchived Then
                question = "Restore this item?"
            Else
                question = "Archive this item?"
            End If

            If MessageBox.Show(
                question,
                "Confirm",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            ) <> DialogResult.Yes Then
                Return
            End If

            Try
                MenuRepository.ArchiveProduct(p.Id, Not p.IsArchived)
                LoadMenuGrid()
                ReloadOrderMenu()
            Catch ex As Exception
                MessageBox.Show("Menu item update failed: " & ex.Message)
            End Try
        End Sub

        '=========================================================
        ' NEW ORDER
        '=========================================================

        Private Function BuildOrderEntry() As Control
            Dim root As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 1,
                .Padding = New Padding(8)
            }

            root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 62))
            root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 38))

            Dim left As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 4
            }

            left.RowStyles.Add(New RowStyle(SizeType.Absolute, 46))
            left.RowStyles.Add(New RowStyle(SizeType.Absolute, 46))
            left.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
            left.RowStyles.Add(New RowStyle(SizeType.Absolute, 48))

            Dim choose As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .WrapContents = False,
                .AutoScroll = True
            }

            menuCombo = New ComboBox With {
                .Width = 330,
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .DisplayMember = "ItemName",
                .ValueMember = "Id"
            }

            quantity = New NumericUpDown With {
                .Minimum = 1,
                .Maximum = 99,
                .Width = 70,
                .Value = 1
            }

            Dim add As New Button With {
                .Text = "Add to Cart",
                .Width = 110
            }

            choose.Controls.AddRange(
                New Control() {
                    New Label With {
                        .Text = "Menu item",
                        .AutoSize = True,
                        .Padding = New Padding(0, 7, 0, 0)
                    },
                    menuCombo,
                    New Label With {
                        .Text = "Qty",
                        .AutoSize = True,
                        .Padding = New Padding(0, 7, 0, 0)
                    },
                    quantity,
                    add
                }
            )

            left.Controls.Add(choose, 0, 0)

            Dim info As New Label With {
                .Text = "Select an available item and quantity, then add it to the cart.",
                .AutoSize = True,
                .Dock = DockStyle.Fill
            }

            left.Controls.Add(info, 0, 1)

            cart = New DataTable()
            cart.Columns.Add("MenuItemId", GetType(Integer))
            cart.Columns.Add("ItemName", GetType(String))
            cart.Columns.Add("UnitPrice", GetType(Decimal))
            cart.Columns.Add("Quantity", GetType(Integer))
            cart.Columns.Add("LineTotal", GetType(Decimal))

            cartGrid = New DataGridView With {
                .Dock = DockStyle.Fill,
                .DataSource = cart,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                .BackgroundColor = Color.White,
                .AllowUserToAddRows = False,
                .ReadOnly = True,
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                .MultiSelect = False
            }

            left.Controls.Add(cartGrid, 0, 2)

            Dim cartButtons As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill
            }

            Dim remove As New Button With {
                .Text = "Remove Selected",
                .Width = 140
            }

            Dim clear As New Button With {
                .Text = "Clear Cart",
                .Width = 100
            }

            cartButtons.Controls.AddRange(New Control() {remove, clear})
            left.Controls.Add(cartButtons, 0, 3)

            root.Controls.Add(left, 0, 0)

            Dim right As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 10,
                .Padding = New Padding(12),
                .BackColor = Color.White
            }

            right.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 115))
            right.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))

            For i As Integer = 0 To 8
                right.RowStyles.Add(New RowStyle(SizeType.Absolute, 46))
            Next

            right.RowStyles.Add(New RowStyle(SizeType.Percent, 100))

            customerName = New TextBox With {
                .Dock = DockStyle.Fill
            }

            orderNotes = New TextBox With {
                .Dock = DockStyle.Fill
            }

            discount = New NumericUpDown With {
                .Dock = DockStyle.Fill,
                .Maximum = 1000000D,
                .DecimalPlaces = 2,
                .ThousandsSeparator = True
            }

            amountPaid = New NumericUpDown With {
                .Dock = DockStyle.Fill,
                .Maximum = 1000000D,
                .DecimalPlaces = 2,
                .ThousandsSeparator = True
            }

            paymentMethod = New ComboBox With {
                .Dock = DockStyle.Fill,
                .DropDownStyle = ComboBoxStyle.DropDownList
            }

            paymentMethod.Items.AddRange(
                New Object() {"Cash", "GCash", "Card", "Other"}
            )
            paymentMethod.SelectedIndex = 0

            lblSubtotal = New Label With {
                .Text = "₱0.00",
                .AutoSize = True,
                .Anchor = AnchorStyles.Left,
                .Font = New Font("Segoe UI", 11, FontStyle.Bold)
            }

            lblTotal = New Label With {
                .Text = "₱0.00",
                .AutoSize = True,
                .Anchor = AnchorStyles.Left,
                .Font = New Font("Segoe UI", 14, FontStyle.Bold),
                .ForeColor = Color.FromArgb(95, 54, 35)
            }

            AddFormRow(right, "Customer", customerName, 0)
            AddFormRow(right, "Notes", orderNotes, 1)
            AddFormRow(right, "Subtotal", lblSubtotal, 2)
            AddFormRow(right, "Discount (₱)", discount, 3)
            AddFormRow(right, "Total", lblTotal, 4)
            AddFormRow(right, "Payment", paymentMethod, 5)
            AddFormRow(right, "Amount paid", amountPaid, 6)

            Dim checkout As New Button With {
                .Text = "Confirm Order & Print Receipt",
                .Dock = DockStyle.Fill,
                .Height = 44,
                .BackColor = Color.FromArgb(120, 73, 48),
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat
            }

            right.Controls.Add(checkout, 0, 8)
            right.SetColumnSpan(checkout, 2)

            root.Controls.Add(right, 1, 0)

            AddHandler add.Click, AddressOf AddCartItem

            AddHandler remove.Click,
                Sub(sender As Object, e As EventArgs)
                    If cartGrid.CurrentRow IsNot Nothing AndAlso
                       Not cartGrid.CurrentRow.IsNewRow Then

                        Dim index As Integer = cartGrid.CurrentRow.Index

                        If index >= 0 AndAlso index < cart.Rows.Count Then
                            cart.Rows.RemoveAt(index)
                            UpdateCartTotals()
                        End If
                    End If
                End Sub

            AddHandler clear.Click,
                Sub(sender As Object, e As EventArgs)
                    cart.Rows.Clear()
                    UpdateCartTotals()
                End Sub

            AddHandler discount.ValueChanged, AddressOf UpdateCartTotals

            AddHandler amountPaid.ValueChanged,
                Sub(sender As Object, e As EventArgs)
                    UpdateCartTotals()
                End Sub

            AddHandler checkout.Click, AddressOf CheckoutClicked

            ReloadOrderMenu()

            Return root
        End Function

        Private Sub AddFormRow(
            p As TableLayoutPanel,
            caption As String,
            c As Control,
            row As Integer
        )
            p.Controls.Add(
                New Label With {
                    .Text = caption,
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left
                },
                0,
                row
            )

            p.Controls.Add(c, 1, row)
        End Sub

        Private Sub ReloadOrderMenu()
            If menuCombo Is Nothing Then Return

            Try
                menuCombo.DataSource = Nothing
                menuCombo.DataSource = MenuRepository.GetAvailableMenu()
                menuCombo.DisplayMember = "ItemName"
                menuCombo.ValueMember = "Id"

                If menuCombo.Items.Count > 0 Then
                    menuCombo.SelectedIndex = 0
                End If

            Catch ex As Exception
                MessageBox.Show("Available menu items could not load: " & ex.Message)
            End Try
        End Sub

        Private Sub AddCartItem(sender As Object, e As EventArgs)
            If menuCombo.SelectedItem Is Nothing Then
                MessageBox.Show(
                    "No available menu items. Add or enable items in Menu Management."
                )
                Return
            End If

            If cart Is Nothing Then Return

            Dim row As DataRowView =
                TryCast(menuCombo.SelectedItem, DataRowView)

            If row Is Nothing Then
                MessageBox.Show("The selected menu item is invalid.")
                Return
            End If

            Dim id As Integer = Convert.ToInt32(row("Id"))
            Dim name As String = Convert.ToString(row("ItemName"))
            Dim price As Decimal = Convert.ToDecimal(row("Price"))
            Dim qty As Integer = Convert.ToInt32(quantity.Value)

            If price < 0D Then
                MessageBox.Show("The menu item price cannot be negative.")
                Return
            End If

            Dim existing As DataRow =
                cart.Select(
                    "MenuItemId=" & id.ToString(CultureInfo.InvariantCulture)
                ).FirstOrDefault()

            If existing Is Nothing Then
                cart.Rows.Add(id, name, price, qty, price * qty)
            Else
                Dim newQuantity As Integer =
                    Convert.ToInt32(existing("Quantity")) + qty

                If newQuantity > 999 Then
                    MessageBox.Show("The quantity for an item cannot exceed 999.")
                    Return
                End If

                existing("Quantity") = newQuantity
                existing("LineTotal") = price * newQuantity
            End If

            UpdateCartTotals()
        End Sub

        Private Function CartSubtotal() As Decimal
            Dim total As Decimal = 0D

            If cart Is Nothing Then Return total

            For Each r As DataRow In cart.Rows
                total += Convert.ToDecimal(r("LineTotal"))
            Next

            Return total
        End Function

        Private Sub UpdateCartTotals(
            Optional sender As Object = Nothing,
            Optional e As EventArgs = Nothing
        )
            If lblSubtotal Is Nothing OrElse lblTotal Is Nothing Then Return

            Dim subTotal As Decimal = CartSubtotal()
            Dim disc As Decimal = discount.Value
            Dim total As Decimal = Math.Max(0D, subTotal - disc)

            lblSubtotal.Text =
                subTotal.ToString("C2", CultureInfo.GetCultureInfo("en-PH"))

            lblTotal.Text =
                total.ToString("C2", CultureInfo.GetCultureInfo("en-PH"))
        End Sub

        Private Sub CheckoutClicked(sender As Object, e As EventArgs)
            Try
                If cart Is Nothing OrElse cart.Rows.Count = 0 Then
                    Throw New ArgumentException(
                        "Add at least one item to the cart."
                    )
                End If

                Dim subtotal As Decimal = CartSubtotal()
                Dim discountAmount As Decimal = discount.Value

                If discountAmount > subtotal Then
                    Throw New ArgumentException(
                        "Discount cannot exceed subtotal."
                    )
                End If

                Dim total As Decimal = subtotal - discountAmount

                If amountPaid.Value < total Then
                    Throw New ArgumentException(
                        "Amount paid is less than the order total."
                    )
                End If

                If paymentMethod.SelectedItem Is Nothing Then
                    Throw New ArgumentException(
                        "Select a payment method."
                    )
                End If

                Dim orderNumber As String =
                    OrderRepository.CreateOrder(
                        customerName.Text.Trim(),
                        orderNotes.Text.Trim(),
                        discountAmount,
                        Convert.ToString(paymentMethod.SelectedItem),
                        amountPaid.Value,
                        currentUser.Username,
                        cart
                    )

                Dim receipt As String =
                    BuildReceipt(
                        orderNumber,
                        customerName.Text,
                        cart,
                        subtotal,
                        discountAmount,
                        total,
                        amountPaid.Value,
                        amountPaid.Value - total
                    )

                Using f As New ReceiptForm(receipt)
                    f.ShowDialog(Me)
                End Using

                cart.Clear()
                customerName.Clear()
                orderNotes.Clear()
                discount.Value = 0D
                amountPaid.Value = 0D

                UpdateCartTotals()
                RefreshAll()

                MessageBox.Show(
                    "Order " & orderNumber & " saved successfully.",
                    "Order Saved",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                )

            Catch ex As Exception
                MessageBox.Show(
                    ex.Message,
                    "Order Validation / Save Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                )
            End Try
        End Sub

        Private Function BuildReceipt(
            no As String,
            customer As String,
            items As DataTable,
            subtotal As Decimal,
            disc As Decimal,
            total As Decimal,
            paid As Decimal,
            change As Decimal
        ) As String

            Dim s As New StringBuilder()

            s.AppendLine(Db.GetSetting("SystemName", "HEARTH CAFÉ"))
            s.AppendLine("SALES RECEIPT")
            s.AppendLine(New String("-"c, 40))
            s.AppendLine("Order: " & no)
            s.AppendLine("Date: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm"))

            Dim customerDisplay As String = "Walk-in"

            If Not String.IsNullOrWhiteSpace(customer) Then
                customerDisplay = customer.Trim()
            End If

            s.AppendLine("Customer: " & customerDisplay)
            s.AppendLine(New String("-"c, 40))

            For Each r As DataRow In items.Rows
                s.AppendLine(
                    Convert.ToString(r("ItemName")) &
                    " x" &
                    Convert.ToString(r("Quantity")) &
                    "  " &
                    Convert.ToDecimal(r("LineTotal")).ToString(
                        "C2",
                        CultureInfo.GetCultureInfo("en-PH")
                    )
                )
            Next

            s.AppendLine(New String("-"c, 40))
            s.AppendLine(
                "Subtotal: " &
                subtotal.ToString("C2", CultureInfo.GetCultureInfo("en-PH"))
            )
            s.AppendLine(
                "Discount: " &
                disc.ToString("C2", CultureInfo.GetCultureInfo("en-PH"))
            )
            s.AppendLine(
                "TOTAL: " &
                total.ToString("C2", CultureInfo.GetCultureInfo("en-PH"))
            )
            s.AppendLine(
                "Paid: " &
                paid.ToString("C2", CultureInfo.GetCultureInfo("en-PH"))
            )
            s.AppendLine(
                "Change: " &
                change.ToString("C2", CultureInfo.GetCultureInfo("en-PH"))
            )
            s.AppendLine()
            s.AppendLine(Db.GetSetting("ReceiptFooter", "Thank you!"))

            Return s.ToString()
        End Function

        '=========================================================
        ' ORDER MANAGEMENT
        '=========================================================

        Private Function BuildOrders() As Control
            Dim root As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .RowCount = 2,
                .ColumnCount = 1,
                .Padding = New Padding(8)
            }

            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 48))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100))

            Dim bar As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .WrapContents = False,
                .AutoScroll = True
            }

            orderSearch = New TextBox With {
                .Width = 260,
                .PlaceholderText = "Search order or customer..."
            }

            Dim search As New Button With {.Text = "Search", .Width = 80}
            Dim refresh As New Button With {.Text = "Refresh", .Width = 80}
            Dim preparing As New Button With {.Text = "Mark Preparing", .Width = 120}
            Dim complete As New Button With {.Text = "Complete", .Width = 90}
            Dim cancel As New Button With {.Text = "Cancel Order", .Width = 100}
            Dim archive As New Button With {.Text = "Archive", .Width = 85}
            Dim details As New Button With {.Text = "View Items", .Width = 90}

            bar.Controls.AddRange(
                New Control() {
                    orderSearch, search, refresh, preparing,
                    complete, cancel, archive, details
                }
            )

            orderGrid = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                .MultiSelect = False,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                .AllowUserToAddRows = False,
                .BackgroundColor = Color.White
            }

            root.Controls.Add(bar, 0, 0)
            root.Controls.Add(orderGrid, 0, 1)

            AddHandler search.Click,
                Sub(sender As Object, e As EventArgs)
                    LoadOrders()
                End Sub

            AddHandler refresh.Click,
                Sub(sender As Object, e As EventArgs)
                    LoadOrders()
                End Sub

            AddHandler orderSearch.KeyDown,
                Sub(sender As Object, e As KeyEventArgs)
                    If e.KeyCode = Keys.Enter Then
                        LoadOrders()
                        e.SuppressKeyPress = True
                    End If
                End Sub

            AddHandler preparing.Click,
                Sub(sender As Object, e As EventArgs)
                    ChangeOrderStatus("Preparing")
                End Sub

            AddHandler complete.Click,
                Sub(sender As Object, e As EventArgs)
                    ChangeOrderStatus("Completed")
                End Sub

            AddHandler cancel.Click,
                Sub(sender As Object, e As EventArgs)
                    ChangeOrderStatus("Cancelled")
                End Sub

            AddHandler archive.Click, AddressOf ArchiveOrder
            AddHandler details.Click, AddressOf ViewOrderItems

            AddHandler orderGrid.CellDoubleClick,
                Sub(sender As Object, e As DataGridViewCellEventArgs)
                    If e.RowIndex >= 0 Then
                        ViewOrderItems(sender, EventArgs.Empty)
                    End If
                End Sub

            Return root
        End Function

        Private Function SelectedOrderId() As Integer
            If orderGrid Is Nothing OrElse orderGrid.CurrentRow Is Nothing Then
                Return 0
            End If

            If Not orderGrid.Columns.Contains("Id") Then Return 0

            Dim value As Object = orderGrid.CurrentRow.Cells("Id").Value

            If value Is Nothing OrElse IsDBNull(value) Then Return 0

            Return Convert.ToInt32(value)
        End Function

        Private Sub LoadOrders()
            If orderGrid Is Nothing Then Return

            Try
                orderGrid.DataSource =
                    OrderRepository.GetOrders(orderSearch.Text.Trim())

                If orderGrid.Columns.Contains("Id") Then
                    orderGrid.Columns("Id").Visible = False
                End If

                Dim moneyColumns As String() = {
                    "Subtotal", "Discount", "Total", "AmountPaid", "ChangeAmount"
                }

                For Each columnName As String In moneyColumns
                    If orderGrid.Columns.Contains(columnName) Then
                        orderGrid.Columns(columnName).DefaultCellStyle.Format = "C2"
                    End If
                Next

            Catch ex As Exception
                MessageBox.Show(
                    "Orders could not load: " & ex.Message,
                    "Order Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                )
            End Try
        End Sub

        Private Sub ChangeOrderStatus(status As String)
            Dim id As Integer = SelectedOrderId()

            If id = 0 Then
                MessageBox.Show("Select an order first.")
                Return
            End If

            If status = "Cancelled" AndAlso Not IsAdmin() Then
                Denied()
                Return
            End If

            Try
                OrderRepository.SetStatus(id, status)
                LoadOrders()

            Catch ex As Exception
                MessageBox.Show(
                    "Status update failed: " & ex.Message,
                    "Order Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                )
            End Try
        End Sub

        Private Sub ArchiveOrder(sender As Object, e As EventArgs)
            If Not IsAdmin() Then
                Denied()
                Return
            End If

            Dim id As Integer = SelectedOrderId()

            If id = 0 Then
                MessageBox.Show("Select an order first.")
                Return
            End If

            If MessageBox.Show(
                "Archive the selected order?",
                "Confirm",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            ) <> DialogResult.Yes Then
                Return
            End If

            Try
                OrderRepository.Archive(id)
                LoadOrders()
            Catch ex As Exception
                MessageBox.Show("Order could not be archived: " & ex.Message)
            End Try
        End Sub

        Private Sub ViewOrderItems(sender As Object, e As EventArgs)
            Dim id As Integer = SelectedOrderId()

            If id = 0 Then Return

            Try
                Dim t As DataTable = OrderRepository.GetOrderItems(id)
                Dim lines As New StringBuilder()

                For Each r As DataRow In t.Rows
                    lines.AppendLine(
                        Convert.ToString(r("ItemName")) &
                        " x" &
                        Convert.ToString(r("Quantity")) &
                        " — " &
                        Convert.ToDecimal(r("LineTotal")).ToString(
                            "C2",
                            CultureInfo.GetCultureInfo("en-PH")
                        )
                    )
                Next

                Dim message As String

                If lines.Length = 0 Then
                    message = "No items found."
                Else
                    message = lines.ToString()
                End If

                MessageBox.Show(
                    message,
                    "Order Items",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                )

            Catch ex As Exception
                MessageBox.Show(
                    "Order items could not load: " & ex.Message,
                    "Order Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                )
            End Try
        End Sub

        '=========================================================
        ' REPORTS
        '=========================================================

        Private Function BuildReports() As Control
            Dim root As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .RowCount = 3,
                .ColumnCount = 1,
                .Padding = New Padding(8)
            }

            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 52))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 48))

            Dim bar As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill,
                .WrapContents = False,
                .AutoScroll = True
            }

            fromDate = New DateTimePicker With {
                .Width = 140,
                .Format = DateTimePickerFormat.Short,
                .Value = DateTime.Today.AddDays(-7)
            }

            toDate = New DateTimePicker With {
                .Width = 140,
                .Format = DateTimePickerFormat.Short,
                .Value = DateTime.Today
            }

            Dim run As New Button With {
                .Text = "Generate Report",
                .Width = 130
            }

            Dim export As New Button With {
                .Text = "Export CSV",
                .Width = 100
            }

            Dim print As New Button With {
                .Text = "Print Report",
                .Width = 100
            }

            bar.Controls.AddRange(
                New Control() {
                    New Label With {
                        .Text = "From",
                        .AutoSize = True,
                        .Padding = New Padding(0, 7, 0, 0)
                    },
                    fromDate,
                    New Label With {
                        .Text = "To",
                        .AutoSize = True,
                        .Padding = New Padding(0, 7, 0, 0)
                    },
                    toDate,
                    run,
                    export,
                    print
                }
            )

            reportGrid = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                .BackgroundColor = Color.White,
                .AllowUserToAddRows = False
            }

            reportSummary = New Label With {
                .Dock = DockStyle.Fill,
                .Font = New Font("Segoe UI", 11, FontStyle.Bold),
                .TextAlign = ContentAlignment.MiddleLeft
            }

            root.Controls.Add(bar, 0, 0)
            root.Controls.Add(reportGrid, 0, 1)
            root.Controls.Add(reportSummary, 0, 2)

            AddHandler run.Click, AddressOf GenerateReport
            AddHandler export.Click, AddressOf ExportReport

            AddHandler print.Click,
                Sub(sender As Object, e As EventArgs)
                    MessageBox.Show(
                        "Use Export CSV to save the report. " &
                        "Open the exported file in Excel and print it.",
                        "Print Report",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    )
                End Sub

            Return root
        End Function

        Private Sub GenerateReport(sender As Object, e As EventArgs)
            If toDate.Value.Date < fromDate.Value.Date Then
                MessageBox.Show("End date must be on or after start date.")
                Return
            End If

            Try
                Dim t As DataTable =
                    OrderRepository.GetReport(
                        fromDate.Value,
                        toDate.Value
                    )

                reportGrid.DataSource = t

                Dim sales As Decimal = 0D
                Dim count As Integer = 0

                For Each r As DataRow In t.Rows
                    If Convert.ToString(r("OrderStatus")) <> "Cancelled" Then
                        If Not IsDBNull(r("Total")) Then
                            sales += Convert.ToDecimal(r("Total"))
                        End If
                    End If

                    count += 1
                Next

                reportSummary.Text =
                    "Transactions: " & count.ToString() &
                    "     Net sales (excluding cancelled): " &
                    sales.ToString(
                        "C2",
                        CultureInfo.GetCultureInfo("en-PH")
                    )

                Dim moneyColumns As String() = {
                    "Subtotal", "Discount", "Total", "AmountPaid", "ChangeAmount"
                }

                For Each columnName As String In moneyColumns
                    If reportGrid.Columns.Contains(columnName) Then
                        reportGrid.Columns(columnName).DefaultCellStyle.Format = "C2"
                    End If
                Next

            Catch ex As Exception
                MessageBox.Show(
                    "Report generation failed: " & ex.Message,
                    "Report Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                )
            End Try
        End Sub

        Private Sub ExportReport(sender As Object, e As EventArgs)
            If reportGrid Is Nothing OrElse reportGrid.DataSource Is Nothing Then
                MessageBox.Show("Generate a report first.")
                Return
            End If

            Using sfd As New SaveFileDialog With {
                .Filter = "CSV files (*.csv)|*.csv",
                .FileName = "HearthCafe-Sales-Report.csv"
            }

                If sfd.ShowDialog(Me) <> DialogResult.OK Then Return

                Try
                    Using sw As New StreamWriter(
                        sfd.FileName,
                        False,
                        New UTF8Encoding(True)
                    )

                        Dim visibleColumns As DataGridViewColumn() =
                            reportGrid.Columns.
                            Cast(Of DataGridViewColumn)().
                            Where(Function(c) c.Visible).
                            ToArray()

                        Dim headers As IEnumerable(Of String) =
                            visibleColumns.Select(
                                Function(c) Csv(c.HeaderText)
                            )

                        sw.WriteLine(String.Join(",", headers))

                        For Each row As DataGridViewRow In reportGrid.Rows
                            If row.IsNewRow Then Continue For

                            Dim values As IEnumerable(Of String) =
                                visibleColumns.Select(
                                    Function(c)
                                        Return Csv(
                                            Convert.ToString(
                                                row.Cells(c.Index).FormattedValue
                                            )
                                        )
                                    End Function
                                )

                            sw.WriteLine(String.Join(",", values))
                        Next

                    End Using

                    MessageBox.Show(
                        "Report exported successfully.",
                        "Export Complete",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    )

                Catch ex As Exception
                    MessageBox.Show(
                        "Export failed: " & ex.Message,
                        "Export Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    )
                End Try
            End Using
        End Sub

        Private Function Csv(value As String) As String
            If value Is Nothing Then value = ""

            Return ChrW(34) &
                   value.Replace(ChrW(34), ChrW(34) & ChrW(34)) &
                   ChrW(34)
        End Function

        '=========================================================
        ' USER ACCOUNTS
        '=========================================================

        Private Function BuildUsers() As Control
            Dim root As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .RowCount = 2,
                .ColumnCount = 1,
                .Padding = New Padding(8)
            }

            root.RowStyles.Add(New RowStyle(SizeType.Absolute, 48))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100))

            Dim bar As New FlowLayoutPanel With {
                .Dock = DockStyle.Fill
            }

            Dim add As New Button With {.Text = "Add User", .Width = 100}
            Dim toggle As New Button With {.Text = "Enable / Disable", .Width = 140}
            Dim reset As New Button With {.Text = "Reset Password", .Width = 130}
            Dim refresh As New Button With {.Text = "Refresh", .Width = 90}

            bar.Controls.AddRange(New Control() {add, toggle, reset, refresh})

            usersGrid = New DataGridView With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                .MultiSelect = False,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                .AllowUserToAddRows = False,
                .BackgroundColor = Color.White
            }

            root.Controls.Add(bar, 0, 0)
            root.Controls.Add(usersGrid, 0, 1)

            AddHandler add.Click, AddressOf AddUserClicked
            AddHandler toggle.Click, AddressOf ToggleUser
            AddHandler reset.Click, AddressOf ResetUserPassword

            AddHandler refresh.Click,
                Sub(sender As Object, e As EventArgs)
                    LoadUsers()
                End Sub

            Return root
        End Function

        Private Sub LoadUsers()
            If usersGrid Is Nothing Then Return

            Try
                usersGrid.DataSource = UserRepository.GetUsers()

                If usersGrid.Columns.Contains("Id") Then
                    usersGrid.Columns("Id").Visible = False
                End If

            Catch ex As Exception
                MessageBox.Show("User accounts could not load: " & ex.Message)
            End Try
        End Sub

        Private Sub AddUserClicked(sender As Object, e As EventArgs)
            If Not IsAdmin() Then
                Denied()
                Return
            End If

            Using f As New NewUserForm()
                If f.ShowDialog(Me) = DialogResult.OK Then
                    LoadUsers()
                End If
            End Using
        End Sub

        Private Function SelectedUserId() As Integer
            If usersGrid Is Nothing OrElse usersGrid.CurrentRow Is Nothing Then
                Return 0
            End If

            If Not usersGrid.Columns.Contains("Id") Then Return 0

            Dim value As Object = usersGrid.CurrentRow.Cells("Id").Value

            If value Is Nothing OrElse IsDBNull(value) Then Return 0

            Return Convert.ToInt32(value)
        End Function

        Private Sub ToggleUser(sender As Object, e As EventArgs)
            If Not IsAdmin() Then
                Denied()
                Return
            End If

            Dim id As Integer = SelectedUserId()

            If id = 0 Then
                MessageBox.Show("Select a user first.")
                Return
            End If

            If id = currentUser.Id Then
                MessageBox.Show(
                    "You cannot disable the account currently in use."
                )
                Return
            End If

            Try
                Dim active As Boolean =
                    Convert.ToBoolean(
                        usersGrid.CurrentRow.Cells("IsActive").Value
                    )

                UserRepository.SetUserActive(id, Not active)
                LoadUsers()

            Catch ex As Exception
                MessageBox.Show("User status update failed: " & ex.Message)
            End Try
        End Sub

        Private Sub ResetUserPassword(sender As Object, e As EventArgs)
            If Not IsAdmin() Then
                Denied()
                Return
            End If

            Dim id As Integer = SelectedUserId()

            If id = 0 Then
                MessageBox.Show("Select a user first.")
                Return
            End If

            Using f As New PasswordResetForm()
                If f.ShowDialog(Me) = DialogResult.OK Then
                    Try
                        UserRepository.ResetPassword(id, f.NewPassword)

                        MessageBox.Show(
                            "Password reset successfully.",
                            "Password Reset",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                        )

                    Catch ex As Exception
                        MessageBox.Show(
                            "Password reset failed: " & ex.Message
                        )
                    End Try
                End If
            End Using
        End Sub

        '=========================================================
        ' SETTINGS
        '=========================================================

        Private Function BuildSettings() As Control
            Dim p As New TableLayoutPanel With {
                .Dock = DockStyle.Top,
                .Width = 800,
                .Height = 300,
                .ColumnCount = 2,
                .RowCount = 5,
                .Padding = New Padding(20)
            }

            p.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 150))
            p.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))

            settingName = New TextBox With {
                .Dock = DockStyle.Fill,
                .Text = Db.GetSetting("SystemName", "HEARTH CAFÉ")
            }

            settingCurrency = New TextBox With {
                .Dock = DockStyle.Fill,
                .Text = Db.GetSetting("Currency", "₱")
            }

            settingFooter = New TextBox With {
                .Dock = DockStyle.Fill,
                .Text = Db.GetSetting("ReceiptFooter", "Thank you!"),
                .Multiline = True,
                .Height = 65,
                .ScrollBars = ScrollBars.Vertical
            }

            p.Controls.Add(
                New Label With {
                    .Text = "System name",
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left
                },
                0,
                0
            )
            p.Controls.Add(settingName, 1, 0)

            p.Controls.Add(
                New Label With {
                    .Text = "Currency symbol",
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left
                },
                0,
                1
            )
            p.Controls.Add(settingCurrency, 1, 1)

            p.Controls.Add(
                New Label With {
                    .Text = "Receipt footer",
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left
                },
                0,
                2
            )
            p.Controls.Add(settingFooter, 1, 2)

            Dim save As New Button With {
                .Text = "Save Settings",
                .Width = 130,
                .Height = 36
            }

            p.Controls.Add(save, 1, 3)

            AddHandler save.Click,
                Sub(sender As Object, e As EventArgs)
                    If String.IsNullOrWhiteSpace(settingName.Text) OrElse
                       String.IsNullOrWhiteSpace(settingCurrency.Text) Then

                        MessageBox.Show(
                            "System name and currency are required."
                        )
                        Return
                    End If

                    Try
                        Db.SaveSetting("SystemName", settingName.Text.Trim())
                        Db.SaveSetting("Currency", settingCurrency.Text.Trim())
                        Db.SaveSetting("ReceiptFooter", settingFooter.Text.Trim())

                        MessageBox.Show(
                            "Settings saved. Restart the application to refresh the title/header."
                        )

                    Catch ex As Exception
                        MessageBox.Show(
                            "Settings could not be saved: " & ex.Message
                        )
                    End Try
                End Sub

            Return p
        End Function

        '=========================================================
        ' REFRESH AND PERMISSIONS
        '=========================================================

        Private Sub RefreshAll()
            LoadMenuGrid()
            ReloadOrderMenu()
            LoadOrders()

            If usersGrid IsNot Nothing Then
                LoadUsers()
            End If
        End Sub

        Private Sub RefreshCurrentTab()
            If tabs.SelectedTab Is Nothing Then Return

            Select Case tabs.SelectedTab.Text
                Case "Menu Management"
                    LoadMenuGrid()
                    ReloadOrderMenu()

                Case "New Order"
                    ReloadOrderMenu()

                Case "Order Management"
                    LoadOrders()

                Case "User Accounts"
                    LoadUsers()
            End Select
        End Sub

        Private Function IsAdmin() As Boolean
            Return String.Equals(
                currentUser.Role,
                "ADMIN",
                StringComparison.OrdinalIgnoreCase
            )
        End Function

        Private Sub Denied()
            MessageBox.Show(
                "Your account does not have permission to perform this action. Contact an administrator.",
                "Access Denied",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
        End Sub

    End Class

    '=============================================================
    ' RECEIPT FORM
    '=============================================================

    Friend Class ReceiptForm
        Inherits Form

        Private ReadOnly receiptText As TextBox

        Public Sub New(receipt As String)
            Text = "Order Receipt"
            StartPosition = FormStartPosition.CenterParent
            ClientSize = New Size(480, 620)

            receiptText = New TextBox With {
                .Dock = DockStyle.Fill,
                .Multiline = True,
                .ReadOnly = True,
                .ScrollBars = ScrollBars.Vertical,
                .Font = New Font("Consolas", 10),
                .Text = If(receipt, "")
            }

            Dim bar As New FlowLayoutPanel With {
                .Dock = DockStyle.Bottom,
                .Height = 48,
                .FlowDirection = FlowDirection.RightToLeft
            }

            Dim save As New Button With {
                .Text = "Save Receipt",
                .Width = 110
            }

            Dim closeButton As New Button With {
                .Text = "Close",
                .Width = 80
            }

            bar.Controls.Add(save)
            bar.Controls.Add(closeButton)

            Controls.Add(receiptText)
            Controls.Add(bar)

            AddHandler closeButton.Click,
                Sub(sender As Object, e As EventArgs)
                    Close()
                End Sub

            AddHandler save.Click,
                Sub(sender As Object, e As EventArgs)
                    Using sfd As New SaveFileDialog With {
                        .Filter = "Text receipt (*.txt)|*.txt",
                        .FileName = "HearthCafe-Receipt.txt"
                    }

                        If sfd.ShowDialog(Me) = DialogResult.OK Then
                            Try
                                File.WriteAllText(sfd.FileName, receiptText.Text)

                                MessageBox.Show(
                                    "Receipt saved successfully.",
                                    "Receipt",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information
                                )

                            Catch ex As Exception
                                MessageBox.Show(
                                    "Receipt could not be saved: " & ex.Message
                                )
                            End Try
                        End If
                    End Using
                End Sub
        End Sub
    End Class

    '=============================================================
    ' NEW USER FORM
    '=============================================================

    Friend Class NewUserForm
        Inherits Form

        Private ReadOnly username As New TextBox With {
            .Dock = DockStyle.Fill
        }

        Private ReadOnly display As New TextBox With {
            .Dock = DockStyle.Fill
        }

        Private ReadOnly role As New ComboBox With {
            .Dock = DockStyle.Fill,
            .DropDownStyle = ComboBoxStyle.DropDownList
        }

        Private ReadOnly password As New TextBox With {
            .Dock = DockStyle.Fill,
            .UseSystemPasswordChar = True
        }

        Private ReadOnly confirm As New TextBox With {
            .Dock = DockStyle.Fill,
            .UseSystemPasswordChar = True
        }

        Public Sub New()
            Text = "Add User Account"
            StartPosition = FormStartPosition.CenterParent
            ClientSize = New Size(470, 340)

            Dim p As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(18),
                .ColumnCount = 2,
                .RowCount = 6
            }

            p.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 130))
            p.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))

            role.Items.AddRange(New Object() {"STAFF", "ADMIN"})
            role.SelectedIndex = 0

            AddRow(p, "Username", username, 0)
            AddRow(p, "Display name", display, 1)
            AddRow(p, "Role", role, 2)
            AddRow(p, "Password", password, 3)
            AddRow(p, "Confirm", confirm, 4)

            Dim save As New Button With {
                .Text = "Create Account",
                .Dock = DockStyle.Fill
            }

            p.RowStyles.Add(New RowStyle(SizeType.Absolute, 48))
            p.Controls.Add(save, 1, 5)

            Controls.Add(p)

            AddHandler save.Click, AddressOf SaveClicked
        End Sub

        Private Sub AddRow(
            p As TableLayoutPanel,
            caption As String,
            c As Control,
            row As Integer
        )
            p.RowStyles.Add(New RowStyle(SizeType.Absolute, 48))

            p.Controls.Add(
                New Label With {
                    .Text = caption,
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left
                },
                0,
                row
            )

            p.Controls.Add(c, 1, row)
        End Sub

        Private Sub SaveClicked(sender As Object, e As EventArgs)
            If String.IsNullOrWhiteSpace(username.Text) OrElse
               String.IsNullOrWhiteSpace(display.Text) OrElse
               password.Text.Length < 8 Then

                MessageBox.Show(
                    "Complete all fields. Password must contain at least 8 characters."
                )
                Return
            End If

            If password.Text <> confirm.Text Then
                MessageBox.Show("Passwords do not match.")
                Return
            End If

            Try
                UserRepository.AddUser(
                    username.Text.Trim(),
                    display.Text.Trim(),
                    Convert.ToString(role.SelectedItem),
                    password.Text
                )

                DialogResult = DialogResult.OK
                Close()

            Catch ex As Exception
                MessageBox.Show(
                    "Could not create account: " & ex.Message,
                    "Account Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                )
            End Try
        End Sub
    End Class

    '=============================================================
    ' PASSWORD RESET FORM
    '=============================================================

    Friend Class PasswordResetForm
        Inherits Form

        Private ReadOnly p1 As New TextBox With {
            .Dock = DockStyle.Fill,
            .UseSystemPasswordChar = True
        }

        Private ReadOnly p2 As New TextBox With {
            .Dock = DockStyle.Fill,
            .UseSystemPasswordChar = True
        }

        Public ReadOnly Property NewPassword As String
            Get
                Return p1.Text
            End Get
        End Property

        Public Sub New()
            Text = "Reset Password"
            StartPosition = FormStartPosition.CenterParent
            ClientSize = New Size(390, 190)

            Dim p As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(15),
                .ColumnCount = 2,
                .RowCount = 3
            }

            p.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 130))
            p.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))

            p.RowStyles.Add(New RowStyle(SizeType.Absolute, 42))
            p.RowStyles.Add(New RowStyle(SizeType.Absolute, 42))
            p.RowStyles.Add(New RowStyle(SizeType.Absolute, 42))

            p.Controls.Add(
                New Label With {
                    .Text = "New password",
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left
                },
                0,
                0
            )

            p.Controls.Add(p1, 1, 0)

            p.Controls.Add(
                New Label With {
                    .Text = "Confirm",
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left
                },
                0,
                1
            )

            p.Controls.Add(p2, 1, 1)

            Dim save As New Button With {
                .Text = "Reset",
                .Dock = DockStyle.Fill
            }

            p.Controls.Add(save, 1, 2)
            Controls.Add(p)

            AddHandler save.Click,
                Sub(sender As Object, e As EventArgs)
                    If p1.Text.Length < 8 Then
                        MessageBox.Show(
                            "Password must contain at least 8 characters."
                        )
                        Return
                    End If

                    If p1.Text <> p2.Text Then
                        MessageBox.Show("Passwords do not match.")
                        Return
                    End If

                    DialogResult = DialogResult.OK
                    Close()
                End Sub
        End Sub
    End Class

End Namespace