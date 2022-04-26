<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmAnexos
#Region "Código generado por el Diseñador de Windows Forms "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'Llamada necesaria para el Diseñador de Windows Forms.
		InitializeComponent()
	End Sub
	'Form invalida a Dispose para limpiar la lista de componentes.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Requerido por el Diseñador de Windows Forms
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents Label1 As System.Windows.Forms.Label

	'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
	'Se puede modificar mediante el Diseñador de Windows Forms.
	'No lo modifique con el editor de código.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.dtgAdicionales = New System.Windows.Forms.DataGridView()
        Me.IdCliente = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RutCarga = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DvCarga = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Nombre = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Paterno = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Materno = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FechaNacimiento = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Parentesco = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.btnVolver = New System.Windows.Forms.Button()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.dtgAdicionales, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(18, 2)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(595, 22)
        Me.Label1.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(18, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(148, 16)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Planes del Asegurado"
        Me.Label2.Visible = False
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Location = New System.Drawing.Point(21, 40)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(796, 358)
        Me.TabControl1.TabIndex = 5
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage1.Controls.Add(Me.dtgAdicionales)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(788, 329)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Adicionales"
        '
        'dtgAdicionales
        '
        Me.dtgAdicionales.AllowUserToAddRows = False
        Me.dtgAdicionales.AllowUserToDeleteRows = False
        Me.dtgAdicionales.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dtgAdicionales.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.IdCliente, Me.RutCarga, Me.DvCarga, Me.Nombre, Me.Paterno, Me.Materno, Me.FechaNacimiento, Me.Parentesco})
        Me.dtgAdicionales.Location = New System.Drawing.Point(7, 7)
        Me.dtgAdicionales.Name = "dtgAdicionales"
        Me.dtgAdicionales.ReadOnly = True
        Me.dtgAdicionales.Size = New System.Drawing.Size(761, 308)
        Me.dtgAdicionales.TabIndex = 0
        '
        'IdCliente
        '
        Me.IdCliente.DataPropertyName = "Id_Cliente"
        Me.IdCliente.HeaderText = "IdCliente"
        Me.IdCliente.Name = "IdCliente"
        Me.IdCliente.ReadOnly = True
        Me.IdCliente.Width = 80
        '
        'RutCarga
        '
        Me.RutCarga.DataPropertyName = "Rut_Carga"
        Me.RutCarga.HeaderText = "RutCarga"
        Me.RutCarga.Name = "RutCarga"
        Me.RutCarga.ReadOnly = True
        '
        'DvCarga
        '
        Me.DvCarga.DataPropertyName = "Dv_Carga"
        Me.DvCarga.HeaderText = "DvCarga"
        Me.DvCarga.Name = "DvCarga"
        Me.DvCarga.ReadOnly = True
        Me.DvCarga.Width = 70
        '
        'Nombre
        '
        Me.Nombre.DataPropertyName = "Nombre"
        Me.Nombre.HeaderText = "Nombre"
        Me.Nombre.Name = "Nombre"
        Me.Nombre.ReadOnly = True
        Me.Nombre.Width = 300
        '
        'Paterno
        '
        Me.Paterno.DataPropertyName = "Paterno"
        Me.Paterno.HeaderText = "Paterno"
        Me.Paterno.Name = "Paterno"
        Me.Paterno.ReadOnly = True
        Me.Paterno.Visible = False
        '
        'Materno
        '
        Me.Materno.DataPropertyName = "Materno"
        Me.Materno.HeaderText = "Materno"
        Me.Materno.Name = "Materno"
        Me.Materno.ReadOnly = True
        Me.Materno.Visible = False
        '
        'FechaNacimiento
        '
        Me.FechaNacimiento.DataPropertyName = "fecha_Nacimiento"
        Me.FechaNacimiento.HeaderText = "FechaNacimiento"
        Me.FechaNacimiento.Name = "FechaNacimiento"
        Me.FechaNacimiento.ReadOnly = True
        Me.FechaNacimiento.Width = 120
        '
        'Parentesco
        '
        Me.Parentesco.DataPropertyName = "Parentesco"
        Me.Parentesco.HeaderText = "Parentesco"
        Me.Parentesco.Name = "Parentesco"
        Me.Parentesco.ReadOnly = True
        Me.Parentesco.Visible = False
        '
        'btnVolver
        '
        Me.btnVolver.Location = New System.Drawing.Point(360, 418)
        Me.btnVolver.Name = "btnVolver"
        Me.btnVolver.Size = New System.Drawing.Size(75, 23)
        Me.btnVolver.TabIndex = 6
        Me.btnVolver.Text = "Volver"
        Me.btnVolver.UseVisualStyleBackColor = True
        '
        'frmAnexos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(829, 467)
        Me.Controls.Add(Me.btnVolver)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAnexos"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Valores Plan"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.dtgAdicionales, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents dtgAdicionales As System.Windows.Forms.DataGridView
    Friend WithEvents IdCliente As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RutCarga As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DvCarga As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Nombre As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Paterno As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Materno As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FechaNacimiento As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Parentesco As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnVolver As System.Windows.Forms.Button
#End Region
End Class