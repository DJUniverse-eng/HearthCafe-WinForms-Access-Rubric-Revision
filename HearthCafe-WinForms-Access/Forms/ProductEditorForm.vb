Imports System.Drawing
Imports System.Windows.Forms

Namespace HearthCafe
    Friend Class ProductEditorForm
        Inherits Form
        Private ReadOnly product As MenuProduct
        Private ReadOnly txtName As New TextBox With {.Dock = DockStyle.Fill}
        Private ReadOnly txtCategory As New TextBox With {.Dock = DockStyle.Fill}
        Private ReadOnly numPrice As New NumericUpDown With {.Dock = DockStyle.Left, .Width = 180, .Maximum = 1000000D, .DecimalPlaces = 2, .Minimum = 0.01D}
        Private ReadOnly txtDescription As New TextBox With {.Dock = DockStyle.Fill, .Multiline = True, .ScrollBars = ScrollBars.Vertical}
        Private ReadOnly txtImage As New TextBox With {.Dock = DockStyle.Fill}
        Private ReadOnly chkAvailable As New CheckBox With {.Text = "Available for ordering", .Checked = True, .AutoSize = True}
        Public Sub New(Optional existing As MenuProduct = Nothing)
            product = If(existing, New MenuProduct())
            Text = If(product.Id = 0, "Add Menu Item", "Edit Menu Item") : StartPosition = FormStartPosition.CenterParent : ClientSize = New Size(570, 440) : MinimumSize = New Size(570, 440)
            Dim p As New TableLayoutPanel With {.Dock = DockStyle.Fill, .Padding = New Padding(18), .ColumnCount = 2, .RowCount = 7}
            p.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 130)) : p.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            For i = 0 To 5 : p.RowStyles.Add(New RowStyle(SizeType.Absolute, If(i = 3, 100, 42))) : Next
            p.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
            AddField(p, "Item name *", txtName, 0) : AddField(p, "Category *", txtCategory, 1) : AddField(p, "Price (₱) *", numPrice, 2)
            AddField(p, "Description", txtDescription, 3) : AddField(p, "Image path", txtImage, 4) : p.Controls.Add(chkAvailable, 1, 5)
            Dim buttons As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .FlowDirection = FlowDirection.RightToLeft}
            Dim save As New Button With {.Text = "Save Item", .Width = 110, .Height = 34} : Dim cancel As New Button With {.Text = "Cancel", .Width = 90, .Height = 34, .DialogResult = DialogResult.Cancel}
            buttons.Controls.Add(save) : buttons.Controls.Add(cancel) : p.Controls.Add(buttons, 1, 6) : Controls.Add(p) : AcceptButton = save : CancelButton = cancel
            txtName.Text = product.Name : txtCategory.Text = product.Category : If product.Price > 0 Then numPrice.Value = Math.Min(numPrice.Maximum, product.Price)
            txtDescription.Text = product.Description : txtImage.Text = product.ImagePath : chkAvailable.Checked = product.IsAvailable
            AddHandler save.Click, AddressOf SaveClicked
        End Sub
        Private Sub AddField(p As TableLayoutPanel, caption As String, c As Control, row As Integer)
            p.Controls.Add(New Label With {.Text = caption, .AutoSize = True, .Anchor = AnchorStyles.Left}, 0, row) : p.Controls.Add(c, 1, row)
        End Sub
        Private Sub SaveClicked(sender As Object, e As EventArgs)
            product.Name = txtName.Text.Trim() : product.Category = txtCategory.Text.Trim() : product.Price = numPrice.Value
            product.Description = txtDescription.Text.Trim() : product.ImagePath = txtImage.Text.Trim() : product.IsAvailable = chkAvailable.Checked
            Try
                MenuRepository.SaveProduct(product) : DialogResult = DialogResult.OK : Close()
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End Sub
    End Class
End Namespace
