
Option Strict On
Option Explicit On
Option Infer On

Imports System
Imports System.Windows.Forms

Namespace HearthCafe

    Friend Module Program

        <STAThread()>
        Public Sub Main()
            Application.EnableVisualStyles()
            Application.SetCompatibleTextRenderingDefault(False)

            Try
                Db.Initialize()

                Using login As New LoginForm()
                    If login.ShowDialog() <> DialogResult.OK Then
                        Return
                    End If

                    If login.AuthenticatedUser Is Nothing Then
                        Return
                    End If

                    Application.Run(New MainForm(login.AuthenticatedUser))
                End Using

            Catch ex As Exception
                MessageBox.Show(
                    "Hearth Café could not start." &
                    Environment.NewLine & ex.Message &
                    Environment.NewLine & Environment.NewLine &
                    "Check the Access Database Engine installation and the database path in Data\Db.vb.",
                    "Startup Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
            End Try
        End Sub

    End Module

End Namespace