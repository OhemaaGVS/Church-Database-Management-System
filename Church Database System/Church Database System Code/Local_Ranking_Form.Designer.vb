<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Local_Ranking_Form
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
        Me.Local_Rank_ListBox = New System.Windows.Forms.ListBox()
        Me.Local_Ranking_Label = New System.Windows.Forms.Label()
        Me.Rank_Label = New System.Windows.Forms.Label()
        Me.Local_Name_ListBox = New System.Windows.Forms.ListBox()
        Me.Local_Number_ListBox = New System.Windows.Forms.ListBox()
        Me.Local_Name = New System.Windows.Forms.Label()
        Me.Number_Label = New System.Windows.Forms.Label()
        Me.ChurchLogoPicBox = New System.Windows.Forms.PictureBox()
        CType(Me.ChurchLogoPicBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Local_Rank_ListBox
        '
        Me.Local_Rank_ListBox.BackColor = System.Drawing.Color.Blue
        Me.Local_Rank_ListBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Local_Rank_ListBox.ForeColor = System.Drawing.SystemColors.Window
        Me.Local_Rank_ListBox.FormattingEnabled = True
        Me.Local_Rank_ListBox.Location = New System.Drawing.Point(22, 113)
        Me.Local_Rank_ListBox.Name = "Local_Rank_ListBox"
        Me.Local_Rank_ListBox.Size = New System.Drawing.Size(120, 420)
        Me.Local_Rank_ListBox.TabIndex = 0
        '
        'Local_Ranking_Label
        '
        Me.Local_Ranking_Label.AutoSize = True
        Me.Local_Ranking_Label.BackColor = System.Drawing.Color.Blue
        Me.Local_Ranking_Label.Font = New System.Drawing.Font("Book Antiqua", 36.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Local_Ranking_Label.ForeColor = System.Drawing.Color.Yellow
        Me.Local_Ranking_Label.Location = New System.Drawing.Point(133, 25)
        Me.Local_Ranking_Label.Name = "Local_Ranking_Label"
        Me.Local_Ranking_Label.Size = New System.Drawing.Size(343, 57)
        Me.Local_Ranking_Label.TabIndex = 1
        Me.Local_Ranking_Label.Text = "Local Ranking"
        '
        'Rank_Label
        '
        Me.Rank_Label.AutoSize = True
        Me.Rank_Label.ForeColor = System.Drawing.Color.Blue
        Me.Rank_Label.Location = New System.Drawing.Point(19, 97)
        Me.Rank_Label.Name = "Rank_Label"
        Me.Rank_Label.Size = New System.Drawing.Size(33, 13)
        Me.Rank_Label.TabIndex = 2
        Me.Rank_Label.Text = "Rank"
        '
        'Local_Name_ListBox
        '
        Me.Local_Name_ListBox.BackColor = System.Drawing.Color.Blue
        Me.Local_Name_ListBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Local_Name_ListBox.ForeColor = System.Drawing.Color.White
        Me.Local_Name_ListBox.FormattingEnabled = True
        Me.Local_Name_ListBox.Location = New System.Drawing.Point(143, 113)
        Me.Local_Name_ListBox.Name = "Local_Name_ListBox"
        Me.Local_Name_ListBox.Size = New System.Drawing.Size(331, 420)
        Me.Local_Name_ListBox.TabIndex = 3
        '
        'Local_Number_ListBox
        '
        Me.Local_Number_ListBox.BackColor = System.Drawing.Color.Blue
        Me.Local_Number_ListBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Local_Number_ListBox.ForeColor = System.Drawing.Color.White
        Me.Local_Number_ListBox.FormattingEnabled = True
        Me.Local_Number_ListBox.Location = New System.Drawing.Point(472, 113)
        Me.Local_Number_ListBox.Name = "Local_Number_ListBox"
        Me.Local_Number_ListBox.Size = New System.Drawing.Size(142, 420)
        Me.Local_Number_ListBox.TabIndex = 4
        '
        'Local_Name
        '
        Me.Local_Name.AutoSize = True
        Me.Local_Name.ForeColor = System.Drawing.Color.Blue
        Me.Local_Name.Location = New System.Drawing.Point(149, 97)
        Me.Local_Name.Name = "Local_Name"
        Me.Local_Name.Size = New System.Drawing.Size(76, 13)
        Me.Local_Name.TabIndex = 5
        Me.Local_Name.Text = "Name of Local"
        '
        'Number_Label
        '
        Me.Number_Label.AutoSize = True
        Me.Number_Label.ForeColor = System.Drawing.Color.Blue
        Me.Number_Label.Location = New System.Drawing.Point(477, 97)
        Me.Number_Label.Name = "Number_Label"
        Me.Number_Label.Size = New System.Drawing.Size(102, 13)
        Me.Number_Label.TabIndex = 6
        Me.Number_Label.Text = "Number of Members"
        '
        'ChurchLogoPicBox
        '
        Me.ChurchLogoPicBox.Image = Global.The_Church_Of_Pentecost_Data_Base_System.My.Resources.Resources.cop
        Me.ChurchLogoPicBox.Location = New System.Drawing.Point(500, 7)
        Me.ChurchLogoPicBox.Name = "ChurchLogoPicBox"
        Me.ChurchLogoPicBox.Size = New System.Drawing.Size(110, 75)
        Me.ChurchLogoPicBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.ChurchLogoPicBox.TabIndex = 18
        Me.ChurchLogoPicBox.TabStop = False
        '
        'Local_Ranking_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(622, 557)
        Me.Controls.Add(Me.ChurchLogoPicBox)
        Me.Controls.Add(Me.Number_Label)
        Me.Controls.Add(Me.Local_Name)
        Me.Controls.Add(Me.Local_Number_ListBox)
        Me.Controls.Add(Me.Local_Name_ListBox)
        Me.Controls.Add(Me.Rank_Label)
        Me.Controls.Add(Me.Local_Ranking_Label)
        Me.Controls.Add(Me.Local_Rank_ListBox)
        Me.Name = "Local_Ranking_Form"
        Me.Text = "Local_Ranking_Form"
        CType(Me.ChurchLogoPicBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Local_Rank_ListBox As System.Windows.Forms.ListBox
    Friend WithEvents Local_Ranking_Label As System.Windows.Forms.Label
    Friend WithEvents Rank_Label As System.Windows.Forms.Label
    Friend WithEvents Local_Name_ListBox As System.Windows.Forms.ListBox
    Friend WithEvents Local_Number_ListBox As System.Windows.Forms.ListBox
    Friend WithEvents Local_Name As System.Windows.Forms.Label
    Friend WithEvents Number_Label As System.Windows.Forms.Label
    Friend WithEvents ChurchLogoPicBox As System.Windows.Forms.PictureBox
End Class
