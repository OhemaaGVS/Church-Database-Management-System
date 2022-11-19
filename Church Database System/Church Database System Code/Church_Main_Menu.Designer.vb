<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Church_Main_Menu
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
        Me.Member_Management_Button = New System.Windows.Forms.Button()
        Me.Local_Management_Button = New System.Windows.Forms.Button()
        Me.Area_Management_Button = New System.Windows.Forms.Button()
        Me.Asset_Management_Button = New System.Windows.Forms.Button()
        Me.District_Management_Button = New System.Windows.Forms.Button()
        Me.Service_Management_Button = New System.Windows.Forms.Button()
        Me.Church_Logo_PicBox = New System.Windows.Forms.PictureBox()
        CType(Me.Church_Logo_PicBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Member_Management_Button
        '
        Me.Member_Management_Button.BackColor = System.Drawing.Color.Blue
        Me.Member_Management_Button.Font = New System.Drawing.Font("Book Antiqua", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Member_Management_Button.ForeColor = System.Drawing.Color.Yellow
        Me.Member_Management_Button.Location = New System.Drawing.Point(116, 216)
        Me.Member_Management_Button.Name = "Member_Management_Button"
        Me.Member_Management_Button.Size = New System.Drawing.Size(237, 38)
        Me.Member_Management_Button.TabIndex = 0
        Me.Member_Management_Button.Text = " Manage Members"
        Me.Member_Management_Button.UseVisualStyleBackColor = False
        '
        'Local_Management_Button
        '
        Me.Local_Management_Button.BackColor = System.Drawing.Color.Blue
        Me.Local_Management_Button.Font = New System.Drawing.Font("Book Antiqua", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Local_Management_Button.ForeColor = System.Drawing.Color.Yellow
        Me.Local_Management_Button.Location = New System.Drawing.Point(116, 275)
        Me.Local_Management_Button.Name = "Local_Management_Button"
        Me.Local_Management_Button.Size = New System.Drawing.Size(237, 43)
        Me.Local_Management_Button.TabIndex = 1
        Me.Local_Management_Button.Text = "Manage Locals"
        Me.Local_Management_Button.UseVisualStyleBackColor = False
        '
        'Area_Management_Button
        '
        Me.Area_Management_Button.BackColor = System.Drawing.Color.Blue
        Me.Area_Management_Button.Font = New System.Drawing.Font("Book Antiqua", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Area_Management_Button.ForeColor = System.Drawing.Color.Yellow
        Me.Area_Management_Button.Location = New System.Drawing.Point(116, 511)
        Me.Area_Management_Button.Name = "Area_Management_Button"
        Me.Area_Management_Button.Size = New System.Drawing.Size(237, 38)
        Me.Area_Management_Button.TabIndex = 2
        Me.Area_Management_Button.Text = "Manage Areas"
        Me.Area_Management_Button.UseVisualStyleBackColor = False
        '
        'Asset_Management_Button
        '
        Me.Asset_Management_Button.BackColor = System.Drawing.Color.Blue
        Me.Asset_Management_Button.Font = New System.Drawing.Font("Book Antiqua", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Asset_Management_Button.ForeColor = System.Drawing.Color.Yellow
        Me.Asset_Management_Button.Location = New System.Drawing.Point(116, 338)
        Me.Asset_Management_Button.Name = "Asset_Management_Button"
        Me.Asset_Management_Button.Size = New System.Drawing.Size(237, 38)
        Me.Asset_Management_Button.TabIndex = 3
        Me.Asset_Management_Button.Text = "Manage Assets"
        Me.Asset_Management_Button.UseVisualStyleBackColor = False
        '
        'District_Management_Button
        '
        Me.District_Management_Button.BackColor = System.Drawing.Color.Blue
        Me.District_Management_Button.Font = New System.Drawing.Font("Book Antiqua", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.District_Management_Button.ForeColor = System.Drawing.Color.Yellow
        Me.District_Management_Button.Location = New System.Drawing.Point(116, 454)
        Me.District_Management_Button.Name = "District_Management_Button"
        Me.District_Management_Button.Size = New System.Drawing.Size(237, 38)
        Me.District_Management_Button.TabIndex = 4
        Me.District_Management_Button.Text = "Manage Districts"
        Me.District_Management_Button.UseVisualStyleBackColor = False
        '
        'Service_Management_Button
        '
        Me.Service_Management_Button.BackColor = System.Drawing.Color.Blue
        Me.Service_Management_Button.Font = New System.Drawing.Font("Book Antiqua", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Service_Management_Button.ForeColor = System.Drawing.Color.Yellow
        Me.Service_Management_Button.Location = New System.Drawing.Point(116, 398)
        Me.Service_Management_Button.Name = "Service_Management_Button"
        Me.Service_Management_Button.Size = New System.Drawing.Size(237, 38)
        Me.Service_Management_Button.TabIndex = 5
        Me.Service_Management_Button.Text = "Manage Services"
        Me.Service_Management_Button.UseVisualStyleBackColor = False
        '
        'Church_Logo_PicBox
        '
        Me.Church_Logo_PicBox.Image = Global.The_Church_Of_Pentecost_Data_Base_System.My.Resources.Resources.cop
        Me.Church_Logo_PicBox.Location = New System.Drawing.Point(104, 12)
        Me.Church_Logo_PicBox.Name = "Church_Logo_PicBox"
        Me.Church_Logo_PicBox.Size = New System.Drawing.Size(271, 185)
        Me.Church_Logo_PicBox.TabIndex = 6
        Me.Church_Logo_PicBox.TabStop = False
        '
        'Church_Main_Menu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(481, 580)
        Me.Controls.Add(Me.Church_Logo_PicBox)
        Me.Controls.Add(Me.Service_Management_Button)
        Me.Controls.Add(Me.District_Management_Button)
        Me.Controls.Add(Me.Asset_Management_Button)
        Me.Controls.Add(Me.Area_Management_Button)
        Me.Controls.Add(Me.Local_Management_Button)
        Me.Controls.Add(Me.Member_Management_Button)
        Me.Name = "Church_Main_Menu"
        Me.Text = "Church Main Menu"
        CType(Me.Church_Logo_PicBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Member_Management_Button As System.Windows.Forms.Button
    Friend WithEvents Local_Management_Button As System.Windows.Forms.Button
    Friend WithEvents Area_Management_Button As System.Windows.Forms.Button
    Friend WithEvents Asset_Management_Button As System.Windows.Forms.Button
    Friend WithEvents District_Management_Button As System.Windows.Forms.Button
    Friend WithEvents Service_Management_Button As System.Windows.Forms.Button
    Friend WithEvents Church_Logo_PicBox As System.Windows.Forms.PictureBox

End Class
