Public Class Form7
    ' Hard-coded admin credentials
    Private ReadOnly adminUsername As String = "admin"
    Private ReadOnly adminPassword As String = "admin123"

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        ' Get the entered username and password
        Dim enteredUsername As String = txtUsername.Text.Trim()
        Dim enteredPassword As String = txtPassword.Text.Trim()

        ' Clear text boxes regardless of login outcome
        txtUsername.Clear()
        txtPassword.Clear()
        txtPassword.PasswordChar = "*"c ' Reset to hidden
        If btnShow.Text = "Hide" Then btnShow.Text = "Show" ' Reset show/hide button

        ' Validate against hard-coded credentials
        If enteredUsername = adminUsername AndAlso enteredPassword = adminPassword Then
            ' Credentials match - show the splash screen then admin form
            SplashScreen1.Show()
            Me.Hide()
        Else
            ' Credentials don't match - show error
            MessageBox.Show("Invalid username or password", "Login Failed",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtUsername.Focus() ' Set focus back to username field
        End If
    End Sub

    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set form properties when it loads
        Me.Text = "Admin Login"
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        ' Initialize password field
        txtPassword.PasswordChar = "*"c
        btnShow.Text = "Show"
    End Sub

    Private Sub btnShow_Click(sender As Object, e As EventArgs) Handles btnShow.Click
        ' Toggle password visibility
        If txtPassword.PasswordChar = "*"c Then
            txtPassword.PasswordChar = Char.MinValue
            btnShow.Text = "Hide"
        Else
            txtPassword.PasswordChar = "*"c
            btnShow.Text = "Show"
        End If
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        ' Clear fields before exiting
        txtUsername.Clear()
        txtPassword.Clear()
        txtPassword.PasswordChar = "*"c
        btnShow.Text = "Show"


        Application.Exit()
    End Sub
End Class