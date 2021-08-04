<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmInsertSymbol
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmInsertSymbol))
        Me.dgvSymbols = New System.Windows.Forms.DataGridView()
        Me.btnInsert = New System.Windows.Forms.Button()
        Me.btnCnacel = New System.Windows.Forms.Button()
        Me.lblSymbol = New System.Windows.Forms.Label()
        Me.fntdlgInsert = New System.Windows.Forms.FontDialog()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cbxSymbolRange = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        CType(Me.dgvSymbols, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvSymbols
        '
        Me.dgvSymbols.AllowUserToAddRows = False
        Me.dgvSymbols.AllowUserToDeleteRows = False
        Me.dgvSymbols.AllowUserToResizeColumns = False
        Me.dgvSymbols.AllowUserToResizeRows = False
        Me.dgvSymbols.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvSymbols.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvSymbols.BackgroundColor = System.Drawing.Color.White
        Me.dgvSymbols.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSymbols.ColumnHeadersVisible = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvSymbols.DefaultCellStyle = DataGridViewCellStyle1
        Me.dgvSymbols.Dock = System.Windows.Forms.DockStyle.Top
        Me.dgvSymbols.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.dgvSymbols.Location = New System.Drawing.Point(0, 0)
        Me.dgvSymbols.MultiSelect = False
        Me.dgvSymbols.Name = "dgvSymbols"
        Me.dgvSymbols.RowHeadersVisible = False
        Me.dgvSymbols.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvSymbols.Size = New System.Drawing.Size(729, 206)
        Me.dgvSymbols.TabIndex = 0
        '
        'btnInsert
        '
        Me.btnInsert.Location = New System.Drawing.Point(267, 272)
        Me.btnInsert.Name = "btnInsert"
        Me.btnInsert.Size = New System.Drawing.Size(75, 23)
        Me.btnInsert.TabIndex = 6
        Me.btnInsert.Text = "Ekle"
        Me.btnInsert.UseVisualStyleBackColor = True
        '
        'btnCnacel
        '
        Me.btnCnacel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCnacel.Location = New System.Drawing.Point(376, 272)
        Me.btnCnacel.Name = "btnCnacel"
        Me.btnCnacel.Size = New System.Drawing.Size(75, 23)
        Me.btnCnacel.TabIndex = 7
        Me.btnCnacel.Text = "İptal"
        Me.btnCnacel.UseVisualStyleBackColor = True
        '
        'lblSymbol
        '
        Me.lblSymbol.AutoSize = True
        Me.lblSymbol.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.lblSymbol.Location = New System.Drawing.Point(197, 244)
        Me.lblSymbol.Name = "lblSymbol"
        Me.lblSymbol.Size = New System.Drawing.Size(54, 18)
        Me.lblSymbol.TabIndex = 4
        Me.lblSymbol.Text = "Simge:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(121, 216)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(103, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Karakter Kod Aralığı:"
        '
        'cbxSymbolRange
        '
        Me.cbxSymbolRange.FormattingEnabled = True
        Me.cbxSymbolRange.Location = New System.Drawing.Point(229, 212)
        Me.cbxSymbolRange.Name = "cbxSymbolRange"
        Me.cbxSymbolRange.Size = New System.Drawing.Size(106, 21)
        Me.cbxSymbolRange.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(364, 216)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(68, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Yazı Tipi Adı:"
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(438, 212)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(177, 21)
        Me.ComboBox1.TabIndex = 9
        '
        'frmInsertSymbol
        '
        Me.AcceptButton = Me.btnInsert
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCnacel
        Me.ClientSize = New System.Drawing.Size(729, 300)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cbxSymbolRange)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblSymbol)
        Me.Controls.Add(Me.btnCnacel)
        Me.Controls.Add(Me.btnInsert)
        Me.Controls.Add(Me.dgvSymbols)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmInsertSymbol"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Simge Ekle"
        CType(Me.dgvSymbols, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

	Friend WithEvents dgvSymbols As DataGridView
	Friend WithEvents btnInsert As Button
    Friend WithEvents btnCnacel As Button
	Friend WithEvents lblSymbol As Label
    Friend WithEvents fntdlgInsert As FontDialog
	Friend WithEvents Label1 As Label
    Friend WithEvents cbxSymbolRange As ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
End Class
