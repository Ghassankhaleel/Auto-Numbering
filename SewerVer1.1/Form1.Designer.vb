<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.Menu1 = New System.Windows.Forms.MenuStrip
        Me.FileItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.LoadShapeFileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator
        Me.ClearToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator
        Me.ExitItem = New System.Windows.Forms.ToolStripMenuItem
        Me.TComboBox1 = New System.Windows.Forms.ToolStripComboBox
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.Button1 = New System.Windows.Forms.Button
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Button7 = New System.Windows.Forms.Button
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Button5 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.ToolStripComboBox1 = New System.Windows.Forms.ToolStripComboBox
        Me.AxMap1 = New AxMapWinGIS.AxMap
        Me.Menu1.SuspendLayout()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxMap1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Menu1
        '
        Me.Menu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileItem, Me.TComboBox1, Me.ToolStripComboBox1})
        Me.Menu1.Location = New System.Drawing.Point(0, 0)
        Me.Menu1.Name = "Menu1"
        Me.Menu1.Size = New System.Drawing.Size(1003, 27)
        Me.Menu1.TabIndex = 0
        Me.Menu1.Text = "MenuStrip1"
        '
        'FileItem
        '
        Me.FileItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSeparator1, Me.LoadShapeFileToolStripMenuItem, Me.ToolStripSeparator4, Me.ClearToolStripMenuItem, Me.ToolStripSeparator3, Me.ExitItem})
        Me.FileItem.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FileItem.Image = Global.SewerVer1._1.My.Resources.Resources.layers1
        Me.FileItem.Name = "FileItem"
        Me.FileItem.Size = New System.Drawing.Size(56, 23)
        Me.FileItem.Text = "&File"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(172, 6)
        '
        'LoadShapeFileToolStripMenuItem
        '
        Me.LoadShapeFileToolStripMenuItem.Name = "LoadShapeFileToolStripMenuItem"
        Me.LoadShapeFileToolStripMenuItem.Size = New System.Drawing.Size(175, 22)
        Me.LoadShapeFileToolStripMenuItem.Text = "Load Shape File"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(172, 6)
        '
        'ClearToolStripMenuItem
        '
        Me.ClearToolStripMenuItem.Name = "ClearToolStripMenuItem"
        Me.ClearToolStripMenuItem.Size = New System.Drawing.Size(175, 22)
        Me.ClearToolStripMenuItem.Text = "Clear"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(172, 6)
        '
        'ExitItem
        '
        Me.ExitItem.Name = "ExitItem"
        Me.ExitItem.Size = New System.Drawing.Size(175, 22)
        Me.ExitItem.Text = "E&xit"
        '
        'TComboBox1
        '
        Me.TComboBox1.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TComboBox1.Name = "TComboBox1"
        Me.TComboBox1.Size = New System.Drawing.Size(121, 23)
        Me.TComboBox1.Text = "Select Manhole No"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Button1
        '
        Me.Button1.Enabled = False
        Me.Button1.Location = New System.Drawing.Point(2, 35)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(84, 38)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Process"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(362, 1)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(100, 22)
        Me.TextBox1.TabIndex = 4
        Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(623, 0)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(283, 23)
        Me.ProgressBar1.TabIndex = 9
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NumericUpDown1.ForeColor = System.Drawing.Color.Blue
        Me.NumericUpDown1.Location = New System.Drawing.Point(-2, 26)
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(84, 21)
        Me.NumericUpDown1.TabIndex = 11
        Me.NumericUpDown1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.NumericUpDown1.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(1, 6)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(80, 17)
        Me.CheckBox1.TabIndex = 12
        Me.CheckBox1.Text = "Initial Value"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.CheckBox1)
        Me.Panel1.Controls.Add(Me.NumericUpDown1)
        Me.Panel1.Location = New System.Drawing.Point(1, 457)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(84, 66)
        Me.Panel1.TabIndex = 13
        '
        'Timer1
        '
        Me.Timer1.Interval = 1
        '
        'Button7
        '
        Me.Button7.Image = Global.SewerVer1._1.My.Resources.Resources.pointer_with_selection
        Me.Button7.Location = New System.Drawing.Point(2, 299)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(84, 35)
        Me.Button7.TabIndex = 16
        Me.Button7.Text = " "
        Me.Button7.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.PictureBox1.Image = Global.SewerVer1._1.My.Resources.Resources.Compass_Rose_clip_art_medium
        Me.PictureBox1.Location = New System.Drawing.Point(0, 355)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(84, 92)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox1.TabIndex = 10
        Me.PictureBox1.TabStop = False
        '
        'Button5
        '
        Me.Button5.Image = Global.SewerVer1._1.My.Resources.Resources.Nov2002_Pan_IM_Button_ben
        Me.Button5.Location = New System.Drawing.Point(2, 249)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(84, 35)
        Me.Button5.TabIndex = 8
        Me.Button5.Text = " "
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Image = Global.SewerVer1._1.My.Resources.Resources.Nov2002_Globe_IM_Button_ben
        Me.Button4.Location = New System.Drawing.Point(2, 196)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(84, 37)
        Me.Button4.TabIndex = 7
        Me.Button4.Text = " "
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Image = Global.SewerVer1._1.My.Resources.Resources.InsetZoomOutOn
        Me.Button3.Location = New System.Drawing.Point(2, 145)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(84, 34)
        Me.Button3.TabIndex = 6
        Me.Button3.Text = " "
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Image = Global.SewerVer1._1.My.Resources.Resources.ZoomInOn
        Me.Button2.Location = New System.Drawing.Point(2, 88)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(84, 38)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = " "
        Me.Button2.UseVisualStyleBackColor = True
        '
        'ToolStripComboBox1
        '
        Me.ToolStripComboBox1.Name = "ToolStripComboBox1"
        Me.ToolStripComboBox1.Size = New System.Drawing.Size(121, 23)
        Me.ToolStripComboBox1.Text = "Select Ground   "
        '
        'AxMap1
        '
        Me.AxMap1.Enabled = True
        Me.AxMap1.Location = New System.Drawing.Point(90, 36)
        Me.AxMap1.Name = "AxMap1"
        Me.AxMap1.OcxState = CType(resources.GetObject("AxMap1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxMap1.Size = New System.Drawing.Size(816, 358)
        Me.AxMap1.TabIndex = 1
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.ClientSize = New System.Drawing.Size(1003, 583)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.AxMap1)
        Me.Controls.Add(Me.Menu1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Menu1.ResumeLayout(False)
        Me.Menu1.PerformLayout()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxMap1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Menu1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents LoadShapeFileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents ToolStripSeparator4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ClearToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AxMap1 As AxMapWinGIS.AxMap
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TComboBox1 As System.Windows.Forms.ToolStripComboBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents NumericUpDown1 As System.Windows.Forms.NumericUpDown
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents ToolStripComboBox1 As System.Windows.Forms.ToolStripComboBox

End Class
