<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
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
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtBoxPath = New System.Windows.Forms.TextBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.CheckedListBox1 = New System.Windows.Forms.CheckedListBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.chkColumn = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.fileNameColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.modifDateColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(370, 298)
        Me.Button1.Margin = New System.Windows.Forms.Padding(1)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(97, 34)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Iniciar"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(17, 22)
        Me.Label1.Margin = New System.Windows.Forms.Padding(1, 0, 1, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(74, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Ruta Archivos"
        '
        'txtBoxPath
        '
        Me.txtBoxPath.Location = New System.Drawing.Point(88, 22)
        Me.txtBoxPath.Margin = New System.Windows.Forms.Padding(1)
        Me.txtBoxPath.Name = "txtBoxPath"
        Me.txtBoxPath.Size = New System.Drawing.Size(301, 20)
        Me.txtBoxPath.TabIndex = 2
        Me.txtBoxPath.Text = "C:\"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(352, 69)
        Me.Button2.Margin = New System.Windows.Forms.Padding(1)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(122, 39)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Seleccionar carpeta"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'CheckedListBox1
        '
        Me.CheckedListBox1.FormattingEnabled = True
        Me.CheckedListBox1.Location = New System.Drawing.Point(6, 30)
        Me.CheckedListBox1.Margin = New System.Windows.Forms.Padding(1)
        Me.CheckedListBox1.Name = "CheckedListBox1"
        Me.CheckedListBox1.Size = New System.Drawing.Size(337, 184)
        Me.CheckedListBox1.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.CheckedListBox1)
        Me.GroupBox1.Location = New System.Drawing.Point(5, 56)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(1)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(1)
        Me.GroupBox1.Size = New System.Drawing.Size(499, 235)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Seleccionar archivo(s) de entrada"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.ImageScalingSize = New System.Drawing.Size(36, 36)
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripStatusLabel2})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 314)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(0, 0, 6, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(887, 22)
        Me.StatusStrip1.TabIndex = 4
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(42, 17)
        Me.ToolStripStatusLabel1.Text = "Status:"
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(55, 17)
        Me.ToolStripStatusLabel2.Text = "Detenido"
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.chkColumn, Me.fileNameColumn, Me.modifDateColumn})
        Me.DataGridView1.Location = New System.Drawing.Point(532, 69)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.Size = New System.Drawing.Size(343, 150)
        Me.DataGridView1.TabIndex = 5
        '
        'chkColumn
        '
        Me.chkColumn.HeaderText = ""
        Me.chkColumn.Name = "chkColumn"
        Me.chkColumn.ReadOnly = True
        Me.chkColumn.Width = 5
        '
        'fileNameColumn
        '
        Me.fileNameColumn.HeaderText = "Archivo"
        Me.fileNameColumn.Name = "fileNameColumn"
        Me.fileNameColumn.ReadOnly = True
        Me.fileNameColumn.Width = 68
        '
        'modifDateColumn
        '
        Me.modifDateColumn.HeaderText = "Ultima fecha modificación"
        Me.modifDateColumn.Name = "modifDateColumn"
        Me.modifDateColumn.ReadOnly = True
        Me.modifDateColumn.Width = 140
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(887, 336)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.txtBoxPath)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button1)
        Me.Margin = New System.Windows.Forms.Padding(1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.GroupBox1.ResumeLayout(False)
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Button2 As Button
    Friend WithEvents CheckedListBox1 As CheckedListBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents txtBoxPath As TextBox
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel2 As ToolStripStatusLabel
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents chkColumn As DataGridViewCheckBoxColumn
    Friend WithEvents fileNameColumn As DataGridViewTextBoxColumn
    Friend WithEvents modifDateColumn As DataGridViewTextBoxColumn
End Class
