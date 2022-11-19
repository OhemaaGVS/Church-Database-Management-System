<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Log_In_Form
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Church_Logo_PicBox = New System.Windows.Forms.PictureBox()
        Me.Log_In_Button = New System.Windows.Forms.Button()
        Me.UserName_Label = New System.Windows.Forms.Label()
        Me.Password_Label = New System.Windows.Forms.Label()
        Me.Username_TextBox = New System.Windows.Forms.TextBox()
        Me.Password_TextBox = New System.Windows.Forms.TextBox()
        Me.Clear_Button = New System.Windows.Forms.Button()
        CType(Me.Church_Logo_PicBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Church_Logo_PicBox
        '
        Me.Church_Logo_PicBox.Image = Global.The_Church_Of_Pentecost_Data_Base_System.My.Resources.Resources.cop
        Me.Church_Logo_PicBox.Location = New System.Drawing.Point(104, 29)
        Me.Church_Logo_PicBox.Name = "Church_Logo_PicBox"
        Me.Church_Logo_PicBox.Size = New System.Drawing.Size(271, 185)
        Me.Church_Logo_PicBox.TabIndex = 7
        Me.Church_Logo_PicBox.TabStop = False
        '
        'Log_In_Button
        '
        Me.Log_In_Button.BackColor = System.Drawing.Color.Blue
        Me.Log_In_Button.Font = New System.Drawing.Font("Book Antiqua", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Log_In_Button.ForeColor = System.Drawing.Color.Yellow
        Me.Log_In_Button.Location = New System.Drawing.Point(97, 345)
        Me.Log_In_Button.Name = "Log_In_Button"
        Me.Log_In_Button.Size = New System.Drawing.Size(97, 32)
        Me.Log_In_Button.TabIndex = 8
        Me.Log_In_Button.Text = "Log In"
        Me.Log_In_Button.UseVisualStyleBackColor = False
        '
        'UserName_Label
        '
        Me.UserName_Label.AutoSize = True
        Me.UserName_Label.BackColor = System.Drawing.Color.Blue
        Me.UserName_Label.Font = New System.Drawing.Font("Book Antiqua", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UserName_Label.ForeColor = System.Drawing.Color.Yellow
        Me.UserName_Label.Location = New System.Drawing.Point(100, 256)
        Me.UserName_Label.Name = "UserName_Label"
        Me.UserName_Label.Size = New System.Drawing.Size(108, 23)
        Me.UserName_Label.TabIndex = 9
        Me.UserName_Label.Text = "User Name"
        '
        'Password_Label
        '
        Me.Password_Label.AutoSize = True
        Me.Password_Label.BackColor = System.Drawing.Color.Blue
        Me.Password_Label.Font = New System.Drawing.Font("Book Antiqua", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Password_Label.ForeColor = System.Drawing.Color.Yellow
        Me.Password_Label.Location = New System.Drawing.Point(100, 300)
        Me.Password_Label.Name = "Password_Label"
        Me.Password_Label.Size = New System.Drawing.Size(94, 23)
        Me.Password_Label.TabIndex = 10
        Me.Password_Label.Text = "Password"
        '
        'Username_TextBox
        '
        Me.Username_TextBox.Location = New System.Drawing.Point(246, 259)
        Me.Username_TextBox.Name = "Username_TextBox"
        Me.Username_TextBox.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.Username_TextBox.Size = New System.Drawing.Size(129, 20)
        Me.Username_TextBox.TabIndex = 11
        '
        'Password_TextBox
        '
        Me.Password_TextBox.Location = New System.Drawing.Point(246, 303)
        Me.Password_TextBox.Name = "Password_TextBox"
        Me.Password_TextBox.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.Password_TextBox.Size = New System.Drawing.Size(129, 20)
        Me.Password_TextBox.TabIndex = 12
        '
        'Clear_Button
        '
        Me.Clear_Button.BackColor = System.Drawing.Color.Blue
        Me.Clear_Button.Font = New System.Drawing.Font("Book Antiqua", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Clear_Button.ForeColor = System.Drawing.Color.Yellow
        Me.Clear_Button.Location = New System.Drawing.Point(246, 345)
        Me.Clear_Button.Name = "Clear_Button"
        Me.Clear_Button.Size = New System.Drawing.Size(97, 32)
        Me.Clear_Button.TabIndex = 13
        Me.Clear_Button.Text = "Clear"
        Me.Clear_Button.UseVisualStyleBackColor = False
        '
        'Log_In_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(489, 438)
        Me.Controls.Add(Me.Clear_Button)
        Me.Controls.Add(Me.Password_TextBox)
        Me.Controls.Add(Me.Username_TextBox)
        Me.Controls.Add(Me.Password_Label)
        Me.Controls.Add(Me.UserName_Label)
        Me.Controls.Add(Me.Log_In_Button)
        Me.Controls.Add(Me.Church_Logo_PicBox)
        Me.Name = "Log_In_Form"
        Me.Text = "Log In"
        CType(Me.Church_Logo_PicBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Church_Logo_PicBox As System.Windows.Forms.PictureBox
    Friend WithEvents Log_In_Button As System.Windows.Forms.Button
    Friend WithEvents UserName_Label As System.Windows.Forms.Label
    Friend WithEvents Password_Label As System.Windows.Forms.Label
    Friend WithEvents Username_TextBox As System.Windows.Forms.TextBox
    Friend WithEvents Password_TextBox As System.Windows.Forms.TextBox
    Friend WithEvents Clear_Button As System.Windows.Forms.Button
End Class
