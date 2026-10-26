Imports System.Drawing
Imports System.Windows.Forms

Namespace HearthCafe
    Public Class LoginForm
        Inherits Form
        Public Property AuthenticatedUser As AppUser
        Private ReadOnly txtUser As New TextBox With {.Width = 260}
        Private ReadOnly txtPassword As New TextBox With {.Width = 260, .UseSystemPasswordChar = True}
        Private ReadOnly btnLogin As New Button With {.Text = "Sign In", .Width = 120, .Height = 36}
        Private ReadOnly btnSetup As New Button With {.Text = "First-time administrator setup", .AutoSize = True}

        Public Sub New()
            Text = "Hearth Café | Sign In" : StartPosition = FormStartPosition.CenterScreen : FormBorderStyle = FormBorderStyle.FixedDialog
            MaximizeBox = False : MinimizeBox = False : ClientSize = New Size(420, 350) : BackColor = Color.FromArgb(249, 244, 236)
            Dim title As New Label With {.Text = "HEARTH CAFÉ", .Font = New Font("Segoe UI", 22, FontStyle.Bold), .ForeColor = Color.FromArgb(95, 54, 35), .AutoSize = True, .Location = New Point(88, 24)}
            Dim subtitle As New Label With {.Text = "Ordering and Management System", .AutoSize = True, .Location = New Point(94, 65)}
            Dim lblU As New Label With {.Text = "Username", .Location = New Point(70, 112), .AutoSize = True}
            txtUser.Location = New Point(70, 135)
            Dim lblP As New Label With {.Text = "Password", .Location = New Point(70, 178), .AutoSize = True}
            txtPassword.Location = New Point(70, 201)
            btnLogin.Location = New Point(70, 250) : btnLogin.BackColor = Color.FromArgb(120, 73, 48) : btnLogin.ForeColor = Color.White : btnLogin.FlatStyle = FlatStyle.Flat
            btnSetup.Location = New Point(70, 300)
            Controls.AddRange(New Control() {title, subtitle, lblU, txtUser, lblP, txtPassword, btnLogin, btnSetup})
            AddHandler btnLogin.Click, AddressOf LoginClicked
            AddHandler btnSetup.Click, AddressOf SetupClicked
            AddHandler txtPassword.KeyDown, Sub(s, e) If e.KeyCode = Keys.Enter Then LoginClicked(s, EventArgs.Empty)
            UpdateSetupButton()
        End Sub

        Private Sub UpdateSetupButton()
            btnSetup.Visible = UserRepository.CountUsers() = 0
        End Sub
        Private Sub SetupClicked(sender As Object, e As EventArgs)
            Using f As New FirstAdminForm()
                If f.ShowDialog(Me) = DialogResult.OK Then UpdateSetupButton()
            End Using
        End Sub
        Private Sub LoginClicked(sender As Object, e As EventArgs)
            Try
                Dim user = UserRepository.Authenticate(txtUser.Text, txtPassword.Text)
                If user Is Nothing Then
                    MessageBox.Show("Invalid username or password, or this account is disabled.", "Sign In", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    txtPassword.Clear() : txtPassword.Focus() : Return
                End If
                AuthenticatedUser = user : DialogResult = DialogResult.OK : Close()
            Catch ex As Exception
                MessageBox.Show("Unable to sign in: " & ex.Message, "Sign In Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
    End Class

    Friend Class FirstAdminForm
        Inherits Form
        Private ReadOnly username As New TextBox With {.Width = 240, .Text = "admin"}
        Private ReadOnly displayName As New TextBox With {.Width = 240, .Text = "System Administrator"}
        Private ReadOnly password As New TextBox With {.Width = 240, .UseSystemPasswordChar = True}
        Private ReadOnly confirm As New TextBox With {.Width = 240, .UseSystemPasswordChar = True}
        Public Sub New()
            Text = "Create First Administrator" : StartPosition = FormStartPosition.CenterParent : FormBorderStyle = FormBorderStyle.FixedDialog : MaximizeBox = False : MinimizeBox = False : ClientSize = New Size(380, 330)
            Dim layout As New TableLayoutPanel With {.Dock = DockStyle.Fill, .Padding = New Padding(18), .ColumnCount = 2, .RowCount = 5}
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120)) : layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            AddRow(layout, "Username", username, 0) : AddRow(layout, "Display name", displayName, 1) : AddRow(layout, "Password", password, 2) : AddRow(layout, "Confirm password", confirm, 3)
            Dim save As New Button With {.Text = "Create Administrator", .Dock = DockStyle.Fill, .Height = 36}
            layout.Controls.Add(save, 1, 4) : Controls.Add(layout) : AddHandler save.Click, AddressOf SaveClicked
        End Sub
        Private Sub AddRow(p As TableLayoutPanel, label As String, c As Control, row As Integer)
            p.RowStyles.Add(New RowStyle(SizeType.Absolute, 52)) : p.Controls.Add(New Label With {.Text = label, .AutoSize = True, .Anchor = AnchorStyles.Left}, 0, row) : p.Controls.Add(c, 1, row)
        End Sub
        Private Sub SaveClicked(sender As Object, e As EventArgs)
            If username.Text.Trim() = "" OrElse displayName.Text.Trim() = "" OrElse password.Text.Length < 8 Then
                MessageBox.Show("Complete all fields. Password must contain at least 8 characters.") : Return
            End If
            If password.Text <> confirm.Text Then MessageBox.Show("Passwords do not match.") : Return
            Try
                If UserRepository.CountUsers() > 0 Then Throw New InvalidOperationException("An administrator already exists. Use normal sign-in.")
                UserRepository.AddUser(username.Text, displayName.Text, "ADMIN", password.Text)
                MessageBox.Show("Administrator account created. Sign in using your new credentials.", "Setup Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
                DialogResult = DialogResult.OK : Close()
            Catch ex As Exception
                MessageBox.Show("Could not create administrator: " & ex.Message)
            End Try
        End Sub
    End Class
End Namespace
