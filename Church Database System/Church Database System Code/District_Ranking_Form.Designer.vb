<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class District_Ranking_Form
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
        Me.Rank_ListBox = New System.Windows.Forms.ListBox()
        Me.District_Ranking_Label = New System.Windows.Forms.Label()
        Me.Rank_Label = New System.Windows.Forms.Label()
        Me.District_ListBox = New System.Windows.Forms.ListBox()
        Me.Number_ListBox = New System.Windows.Forms.ListBox()
        Me.Name_Label = New System.Windows.Forms.Label()
        Me.Number_Label = New System.Windows.Forms.Label()
        Me.ChurchLogoPicBox = New System.Windows.Forms.PictureBox()
        CType(Me.ChurchLogoPicBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Rank_ListBox
        '
        Me.Rank_ListBox.BackColor = System.Drawing.Color.Blue
        Me.Rank_ListBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Rank_ListBox.ForeColor = System.Drawing.SystemColors.Window
        Me.Rank_ListBox.FormattingEnabled = True
        Me.Rank_ListBox.Location = New System.Drawing.Point(22, 113)
        Me.Rank_ListBox.Name = "Rank_ListBox"
        Me.Rank_ListBox.Size = New System.Drawing.Size(120, 420)
        Me.Rank_ListBox.TabIndex = 0
        '
        'District_Ranking_Label
        '
        Me.District_Ranking_Label.AutoSize = True
        Me.District_Ranking_Label.BackColor = System.Drawing.Color.Blue
        Me.District_Ranking_Label.Font = New System.Drawing.Font("Book Antiqua", 36.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.District_Ranking_Label.ForeColor = System.Drawing.Color.Yellow
        Me.District_Ranking_Label.Location = New System.Drawing.Point(37, 40)
        Me.District_Ranking_Label.Name = "District_Ranking_Label"
        Me.District_Ranking_Label.Size = New System.Drawing.Size(391, 57)
        Me.District_Ranking_Label.TabIndex = 1
        Me.District_Ranking_Label.Text = "District Ranking"
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
        'District_ListBox
        '
        Me.District_ListBox.BackColor = System.Drawing.Color.Blue
        Me.District_ListBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.District_ListBox.ForeColor = System.Drawing.Color.White
        Me.District_ListBox.FormattingEnabled = True
        Me.District_ListBox.Location = New System.Drawing.Point(143, 113)
        Me.District_ListBox.Name = "District_ListBox"
        Me.District_ListBox.Size = New System.Drawing.Size(331, 420)
        Me.District_ListBox.TabIndex = 3
        '
        'Number_ListBox
        '
        Me.Number_ListBox.BackColor = System.Drawing.Color.Blue
        Me.Number_ListBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Number_ListBox.ForeColor = System.Drawing.Color.White
        Me.Number_ListBox.FormattingEnabled = True
        Me.Number_ListBox.Location = New System.Drawing.Point(471, 113)
        Me.Number_ListBox.Name = "Number_ListBox"
        Me.Number_ListBox.Size = New System.Drawing.Size(120, 420)
        Me.Number_ListBox.TabIndex = 4
        '
        'Name_Label
        '
        Me.Name_Label.AutoSize = True
        Me.Name_Label.ForeColor = System.Drawing.Color.Blue
        Me.Name_Label.Location = New System.Drawing.Point(149, 97)
        Me.Name_Label.Name = "Name_Label"
        Me.Name_Label.Size = New System.Drawing.Size(82, 13)
        Me.Name_Label.TabIndex = 5
        Me.Name_Label.Text = "Name of District"
        '
        'Number_Label
        '
        Me.Number_Label.AutoSize = True
        Me.Number_Label.ForeColor = System.Drawing.Color.Blue
        Me.Number_Label.Location = New System.Drawing.Point(477, 97)
        Me.Number_Label.Name = "Number_Label"
        Me.Number_Label.Size = New System.Drawing.Size(86, 13)
        Me.Number_Label.TabIndex = 6
        Me.Number_Label.Text = "Number of locals"
        '
        'ChurchLogoPicBox
        '
        Me.ChurchLogoPicBox.Image = Global.The_Church_Of_Pentecost_Data_Base_System.My.Resources.Resources.cop
        Me.ChurchLogoPicBox.Location = New System.Drawing.Point(453, 22)
        Me.ChurchLogoPicBox.Name = "ChurchLogoPicBox"
        Me.ChurchLogoPicBox.Size = New System.Drawing.Size(110, 75)
        Me.ChurchLogoPicBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.ChurchLogoPicBox.TabIndex = 20
        Me.ChurchLogoPicBox.TabStop = False
        '
        'District_Ranking_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(596, 572)
        Me.Controls.Add(Me.ChurchLogoPicBox)
        Me.Controls.Add(Me.Number_Label)
        Me.Controls.Add(Me.Name_Label)
        Me.Controls.Add(Me.Number_ListBox)
        Me.Controls.Add(Me.District_ListBox)
        Me.Controls.Add(Me.Rank_Label)
        Me.Controls.Add(Me.District_Ranking_Label)
        Me.Controls.Add(Me.Rank_ListBox)
        Me.Name = "District_Ranking_Form"
        Me.Text = "District_Ranking_Form"
        CType(Me.ChurchLogoPicBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Rank_ListBox As System.Windows.Forms.ListBox
    Friend WithEvents District_Ranking_Label As System.Windows.Forms.Label
    Friend WithEvents Rank_Label As System.Windows.Forms.Label
    Friend WithEvents District_ListBox As System.Windows.Forms.ListBox
    Friend WithEvents Number_ListBox As System.Windows.Forms.ListBox
    Friend WithEvents Name_Label As System.Windows.Forms.Label
    Friend WithEvents Number_Label As System.Windows.Forms.Label
    Friend WithEvents ChurchLogoPicBox As System.Windows.Forms.PictureBox
End Class
