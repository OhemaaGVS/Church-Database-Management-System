<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Area_Ranking_Form
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
        Me.Area_Ranking_Label = New System.Windows.Forms.Label()
        Me.Rank_Label = New System.Windows.Forms.Label()
        Me.Area_ListBox = New System.Windows.Forms.ListBox()
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
        Me.Rank_ListBox.ForeColor = System.Drawing.Color.White
        Me.Rank_ListBox.FormattingEnabled = True
        Me.Rank_ListBox.Location = New System.Drawing.Point(22, 113)
        Me.Rank_ListBox.Name = "Rank_ListBox"
        Me.Rank_ListBox.Size = New System.Drawing.Size(120, 420)
        Me.Rank_ListBox.TabIndex = 0
        '
        'Area_Ranking_Label
        '
        Me.Area_Ranking_Label.AutoSize = True
        Me.Area_Ranking_Label.BackColor = System.Drawing.Color.Blue
        Me.Area_Ranking_Label.Font = New System.Drawing.Font("Book Antiqua", 36.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Area_Ranking_Label.ForeColor = System.Drawing.Color.Yellow
        Me.Area_Ranking_Label.Location = New System.Drawing.Point(127, 26)
        Me.Area_Ranking_Label.Name = "Area_Ranking_Label"
        Me.Area_Ranking_Label.Size = New System.Drawing.Size(330, 57)
        Me.Area_Ranking_Label.TabIndex = 1
        Me.Area_Ranking_Label.Text = "Area Ranking"
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
        'Area_ListBox
        '
        Me.Area_ListBox.BackColor = System.Drawing.Color.Blue
        Me.Area_ListBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Area_ListBox.ForeColor = System.Drawing.SystemColors.Window
        Me.Area_ListBox.FormattingEnabled = True
        Me.Area_ListBox.Location = New System.Drawing.Point(137, 113)
        Me.Area_ListBox.Name = "Area_ListBox"
        Me.Area_ListBox.Size = New System.Drawing.Size(337, 420)
        Me.Area_ListBox.TabIndex = 3
        '
        'Number_ListBox
        '
        Me.Number_ListBox.BackColor = System.Drawing.Color.Blue
        Me.Number_ListBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Number_ListBox.ForeColor = System.Drawing.Color.White
        Me.Number_ListBox.FormattingEnabled = True
        Me.Number_ListBox.Location = New System.Drawing.Point(473, 113)
        Me.Number_ListBox.Name = "Number_ListBox"
        Me.Number_ListBox.Size = New System.Drawing.Size(127, 420)
        Me.Number_ListBox.TabIndex = 4
        '
        'Name_Label
        '
        Me.Name_Label.AutoSize = True
        Me.Name_Label.ForeColor = System.Drawing.Color.Blue
        Me.Name_Label.Location = New System.Drawing.Point(149, 97)
        Me.Name_Label.Name = "Name_Label"
        Me.Name_Label.Size = New System.Drawing.Size(72, 13)
        Me.Name_Label.TabIndex = 5
        Me.Name_Label.Text = "Name of Area"
        '
        'Number_Label
        '
        Me.Number_Label.AutoSize = True
        Me.Number_Label.ForeColor = System.Drawing.Color.Blue
        Me.Number_Label.Location = New System.Drawing.Point(477, 97)
        Me.Number_Label.Name = "Number_Label"
        Me.Number_Label.Size = New System.Drawing.Size(96, 13)
        Me.Number_Label.TabIndex = 6
        Me.Number_Label.Text = "Number of Districts"
        '
        'ChurchLogoPicBox
        '
        Me.ChurchLogoPicBox.Image = Global.The_Church_Of_Pentecost_Data_Base_System.My.Resources.Resources.cop
        Me.ChurchLogoPicBox.Location = New System.Drawing.Point(473, 19)
        Me.ChurchLogoPicBox.Name = "ChurchLogoPicBox"
        Me.ChurchLogoPicBox.Size = New System.Drawing.Size(110, 75)
        Me.ChurchLogoPicBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.ChurchLogoPicBox.TabIndex = 21
        Me.ChurchLogoPicBox.TabStop = False
        '
        'Area_Ranking_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(600, 563)
        Me.Controls.Add(Me.ChurchLogoPicBox)
        Me.Controls.Add(Me.Number_Label)
        Me.Controls.Add(Me.Name_Label)
        Me.Controls.Add(Me.Number_ListBox)
        Me.Controls.Add(Me.Area_ListBox)
        Me.Controls.Add(Me.Rank_Label)
        Me.Controls.Add(Me.Area_Ranking_Label)
        Me.Controls.Add(Me.Rank_ListBox)
        Me.Name = "Area_Ranking_Form"
        Me.Text = "Area_Ranking_Form"
        CType(Me.ChurchLogoPicBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Rank_ListBox As System.Windows.Forms.ListBox
    Friend WithEvents Area_Ranking_Label As System.Windows.Forms.Label
    Friend WithEvents Rank_Label As System.Windows.Forms.Label
    Friend WithEvents Area_ListBox As System.Windows.Forms.ListBox
    Friend WithEvents Number_ListBox As System.Windows.Forms.ListBox
    Friend WithEvents Name_Label As System.Windows.Forms.Label
    Friend WithEvents Number_Label As System.Windows.Forms.Label
    Friend WithEvents ChurchLogoPicBox As System.Windows.Forms.PictureBox
End Class
