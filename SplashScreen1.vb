Public NotInheritable Class SplashScreen1
    Private WithEvents closeTimer As New Timer()
    Private elapsedSeconds As Integer = 0
    Private Const DISPLAY_TIME As Integer = 5 ' Display for 5 seconds

    Private Sub SplashScreen1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Set up the dialog text at runtime according to the application's assembly information.  
        If My.Application.Info.Title <> "" Then
            ApplicationTitle.Text = My.Application.Info.Title
        Else
            ApplicationTitle.Text = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
        End If

        Version.Text = System.String.Format(Version.Text, My.Application.Info.Version.Major, My.Application.Info.Version.Minor)
        Copyright.Text = My.Application.Info.Copyright

        ' Set up timer to close splash screen after 10 seconds
        closeTimer.Interval = 500 ' 1 second intervals
        closeTimer.Start()

        ' Optional: Add a progress bar or countdown display
        AddCountdownDisplay()
    End Sub

    Private Sub AddCountdownDisplay()
        ' Add a simple countdown label
        Dim countdownLabel As New Label With {
            .Text = DISPLAY_TIME.ToString(),
            .Font = New Font("Arial", 12, FontStyle.Bold),
            .ForeColor = Color.White,
            .BackColor = Color.Transparent,
            .AutoSize = True
        }
        countdownLabel.Location = New Point(Me.Width - 50, 10)
        Me.Controls.Add(countdownLabel)
        countdownLabel.BringToFront()

        ' Update the countdown each second
        AddHandler closeTimer.Tick, Sub()
                                        elapsedSeconds += 1
                                        Dim remaining = DISPLAY_TIME - elapsedSeconds
                                        countdownLabel.Text = remaining.ToString()

                                        If remaining <= 0 Then
                                            CloseSplashScreen()
                                        End If
                                    End Sub
    End Sub

    Private Sub CloseSplashScreen()
        closeTimer.Stop()

        ' Optional: Add fade-out effect
        For fadeOut As Double = 1.0 To 0.0 Step -0.05
            Me.Opacity = fadeOut
            Application.DoEvents()
            Threading.Thread.Sleep(50)
        Next

        Me.Close()

        ' Show the main form (replace Form2 with your actual main form)
        Dim mainForm As New Form2()
        mainForm.Show()
    End Sub

    Private Sub ApplicationTitle_Click(sender As Object, e As EventArgs) Handles ApplicationTitle.Click
        ' Allow users to click to close splash screen early
        CloseSplashScreen()
    End Sub

    ' Make the entire splash screen clickable to close
    Private Sub SplashScreen1_Click(sender As Object, e As EventArgs) Handles Me.Click
        CloseSplashScreen()
    End Sub

    Private Sub MainLayoutPanel_Paint(sender As Object, e As PaintEventArgs) Handles MainLayoutPanel.Paint
        ' Optional: Add custom painting if needed
    End Sub
End Class