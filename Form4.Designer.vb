<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form4
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form4))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboEmployeeID = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtBasicSalary = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtOTHours = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtPAYE = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtSSNIT = New System.Windows.Forms.TextBox()
        Me.txtNHIL = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtGFLvy = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtNetSalary = New System.Windows.Forms.TextBox()
        Me.btnCalculate = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnViewSlips = New System.Windows.Forms.Button()
        Me.txtFirstName = New System.Windows.Forms.TextBox()
        Me.txtSurname = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtPosition = New System.Windows.Forms.TextBox()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.txtOTPay = New System.Windows.Forms.TextBox()
        Me.DTPpaid = New System.Windows.Forms.DateTimePicker()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(8, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(102, 19)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Employee ID:"
        '
        'cboEmployeeID
        '
        Me.cboEmployeeID.FormattingEnabled = True
        Me.cboEmployeeID.Location = New System.Drawing.Point(124, 23)
        Me.cboEmployeeID.Name = "cboEmployeeID"
        Me.cboEmployeeID.Size = New System.Drawing.Size(66, 21)
        Me.cboEmployeeID.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(8, 196)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(98, 19)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Basic Salary:"
        '
        'txtBasicSalary
        '
        Me.txtBasicSalary.Location = New System.Drawing.Point(124, 197)
        Me.txtBasicSalary.Name = "txtBasicSalary"
        Me.txtBasicSalary.Size = New System.Drawing.Size(66, 20)
        Me.txtBasicSalary.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(355, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(122, 19)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Overtime Hours:"
        '
        'txtOTHours
        '
        Me.txtOTHours.Location = New System.Drawing.Point(496, 25)
        Me.txtOTHours.Name = "txtOTHours"
        Me.txtOTHours.Size = New System.Drawing.Size(66, 20)
        Me.txtOTHours.TabIndex = 5
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(396, 181)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(81, 19)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "PAYE Tax:"
        '
        'txtPAYE
        '
        Me.txtPAYE.Location = New System.Drawing.Point(494, 181)
        Me.txtPAYE.Name = "txtPAYE"
        Me.txtPAYE.Size = New System.Drawing.Size(66, 20)
        Me.txtPAYE.TabIndex = 7
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(396, 225)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(59, 19)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "SSNIT:"
        '
        'txtSSNIT
        '
        Me.txtSSNIT.Location = New System.Drawing.Point(494, 224)
        Me.txtSSNIT.Name = "txtSSNIT"
        Me.txtSSNIT.Size = New System.Drawing.Size(66, 20)
        Me.txtSSNIT.TabIndex = 9
        '
        'txtNHIL
        '
        Me.txtNHIL.Location = New System.Drawing.Point(495, 104)
        Me.txtNHIL.Name = "txtNHIL"
        Me.txtNHIL.Size = New System.Drawing.Size(67, 20)
        Me.txtNHIL.TabIndex = 10
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(396, 105)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(54, 19)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "NHIL:"
        '
        'txtGFLvy
        '
        Me.txtGFLvy.Location = New System.Drawing.Point(494, 142)
        Me.txtGFLvy.Name = "txtGFLvy"
        Me.txtGFLvy.Size = New System.Drawing.Size(67, 20)
        Me.txtGFLvy.TabIndex = 13
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(205, 264)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(86, 19)
        Me.Label8.TabIndex = 14
        Me.Label8.Text = "Net Salary:"
        '
        'txtNetSalary
        '
        Me.txtNetSalary.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNetSalary.Location = New System.Drawing.Point(311, 259)
        Me.txtNetSalary.Multiline = True
        Me.txtNetSalary.Name = "txtNetSalary"
        Me.txtNetSalary.Size = New System.Drawing.Size(87, 27)
        Me.txtNetSalary.TabIndex = 15
        '
        'btnCalculate
        '
        Me.btnCalculate.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCalculate.Location = New System.Drawing.Point(66, 352)
        Me.btnCalculate.Name = "btnCalculate"
        Me.btnCalculate.Size = New System.Drawing.Size(93, 40)
        Me.btnCalculate.TabIndex = 16
        Me.btnCalculate.Text = "Calculate"
        Me.btnCalculate.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Location = New System.Drawing.Point(165, 352)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(93, 40)
        Me.btnSave.TabIndex = 17
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnBack
        '
        Me.btnBack.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBack.Location = New System.Drawing.Point(463, 352)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(93, 40)
        Me.btnBack.TabIndex = 18
        Me.btnBack.Text = "Back"
        Me.btnBack.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(360, 143)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(117, 19)
        Me.Label7.TabIndex = 12
        Me.Label7.Text = "GETFund Levy:"
        '
        'btnViewSlips
        '
        Me.btnViewSlips.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnViewSlips.Location = New System.Drawing.Point(264, 352)
        Me.btnViewSlips.Name = "btnViewSlips"
        Me.btnViewSlips.Size = New System.Drawing.Size(93, 40)
        Me.btnViewSlips.TabIndex = 19
        Me.btnViewSlips.Text = "View Slips"
        Me.btnViewSlips.UseVisualStyleBackColor = True
        '
        'txtFirstName
        '
        Me.txtFirstName.Location = New System.Drawing.Point(124, 67)
        Me.txtFirstName.Name = "txtFirstName"
        Me.txtFirstName.Size = New System.Drawing.Size(167, 20)
        Me.txtFirstName.TabIndex = 20
        '
        'txtSurname
        '
        Me.txtSurname.Location = New System.Drawing.Point(124, 106)
        Me.txtSurname.Name = "txtSurname"
        Me.txtSurname.Size = New System.Drawing.Size(167, 20)
        Me.txtSurname.TabIndex = 21
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(8, 66)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(85, 19)
        Me.Label9.TabIndex = 22
        Me.Label9.Text = "FirstName:"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(8, 107)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(73, 19)
        Me.Label10.TabIndex = 23
        Me.Label10.Text = "Surname:"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(8, 143)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(67, 19)
        Me.Label11.TabIndex = 24
        Me.Label11.Text = "Position:"
        '
        'txtPosition
        '
        Me.txtPosition.Location = New System.Drawing.Point(124, 144)
        Me.txtPosition.Name = "txtPosition"
        Me.txtPosition.Size = New System.Drawing.Size(167, 20)
        Me.txtPosition.TabIndex = 25
        '
        'btnClear
        '
        Me.btnClear.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClear.Location = New System.Drawing.Point(364, 352)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(93, 40)
        Me.btnClear.TabIndex = 26
        Me.btnClear.Text = "Clear"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(371, 68)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(106, 19)
        Me.Label12.TabIndex = 27
        Me.Label12.Text = "Overtime Pay:"
        '
        'txtOTPay
        '
        Me.txtOTPay.Location = New System.Drawing.Point(496, 69)
        Me.txtOTPay.Name = "txtOTPay"
        Me.txtOTPay.Size = New System.Drawing.Size(66, 20)
        Me.txtOTPay.TabIndex = 28
        '
        'DTPpaid
        '
        Me.DTPpaid.Location = New System.Drawing.Point(241, 308)
        Me.DTPpaid.Name = "DTPpaid"
        Me.DTPpaid.Size = New System.Drawing.Size(198, 20)
        Me.DTPpaid.TabIndex = 34
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(153, 314)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(67, 13)
        Me.Label13.TabIndex = 35
        Me.Label13.Text = "Date Paid:"
        '
        'Form4
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(616, 404)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.DTPpaid)
        Me.Controls.Add(Me.txtOTPay)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.btnClear)
        Me.Controls.Add(Me.txtPosition)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.txtSurname)
        Me.Controls.Add(Me.txtFirstName)
        Me.Controls.Add(Me.btnViewSlips)
        Me.Controls.Add(Me.btnBack)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnCalculate)
        Me.Controls.Add(Me.txtNetSalary)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.txtGFLvy)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtNHIL)
        Me.Controls.Add(Me.txtSSNIT)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtPAYE)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtOTHours)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtBasicSalary)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cboEmployeeID)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form4"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "PayMaster-Payroll"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents cboEmployeeID As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents txtBasicSalary As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents txtOTHours As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents txtPAYE As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents txtSSNIT As TextBox
    Friend WithEvents txtNHIL As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents txtGFLvy As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents txtNetSalary As TextBox
    Friend WithEvents btnCalculate As Button
    Friend WithEvents btnSave As Button
    Friend WithEvents btnBack As Button
    Friend WithEvents Label7 As Label
    Friend WithEvents btnViewSlips As Button
    Friend WithEvents txtFirstName As TextBox
    Friend WithEvents txtSurname As TextBox
    Friend WithEvents Label9 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents txtPosition As TextBox
    Friend WithEvents btnClear As Button
    Friend WithEvents Label12 As Label
    Friend WithEvents txtOTPay As TextBox
    Friend WithEvents DTPpaid As DateTimePicker
    Friend WithEvents Label13 As Label
End Class
