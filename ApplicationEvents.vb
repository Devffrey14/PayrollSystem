Namespace My
    ' The following events are available for MyApplication:
    ' Startup: Raised when the application starts, before the startup form is created.
    ' Shutdown: Raised after all application forms are closed.  This event is not raised if the application terminates abnormally.
    ' UnhandledException: Raised if the application encounters an unhandled exception.
    ' StartupNextInstance: Raised when launching a single-instance application and the application is already active. 
    ' NetworkAvailabilityChanged: Raised when the network connection is connected or disconnected.
    Partial Friend Class MyApplication
    End Class
    Private Sub CheckAndCreateDatabase()
        Try
            ' For SQLite
            If Not File.Exists("MyDatabase.db") Then
                SQLiteConnection.CreateFile("MyDatabase.db")
                CreateDatabaseSchema()
            End If

            ' For LocalDB alternative:
            ' If Not File.Exists(Path.Combine(Application.StartupPath, "MyDatabase.mdf")) Then
            '     CreateLocalDBDatabase()
            ' End If
        Catch ex As Exception
            MessageBox.Show("Database initialization failed: " & ex.Message)
        End Try
    End Sub

    Private Sub CreateDatabaseSchema()
        Using conn As New SQLiteConnection(connectionString)
            conn.Open()

            Using cmd As New SQLiteCommand(conn)
                cmd.CommandText = "CREATE TABLE IF NOT EXISTS Employees (" &
                                 "EmployeeID INTEGER PRIMARY KEY AUTOINCREMENT, " &
                                 "FirstName TEXT NOT NULL, " &
                                 "Surname TEXT NOT NULL, " &
                                 "Position TEXT, " &
                                 "BasicSalary REAL)"
                cmd.ExecuteNonQuery()

                ' Add more tables as needed
            End Using
        End Using
    End Sub
End Namespace
