<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmBeneficiario
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBeneficiario))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label38 = New System.Windows.Forms.Label()
        Me.dtFechaNacBen = New System.Windows.Forms.DateTimePicker()
        Me.dtBeneficiario = New System.Windows.Forms.DataGridView()
        Me.Label70 = New System.Windows.Forms.Label()
        Me.txtPorcentajeBen = New System.Windows.Forms.TextBox()
        Me.Label71 = New System.Windows.Forms.Label()
        Me.cmdEliminarBen = New System.Windows.Forms.Button()
        Me.cmdModificarBen = New System.Windows.Forms.Button()
        Me.cmdAgregarBen = New System.Windows.Forms.Button()
        Me.Label72 = New System.Windows.Forms.Label()
        Me.Label73 = New System.Windows.Forms.Label()
        Me.cmbParentescoBen = New System.Windows.Forms.ComboBox()
        Me.txtDvBen = New System.Windows.Forms.TextBox()
        Me.Label74 = New System.Windows.Forms.Label()
        Me.txtRutBen = New System.Windows.Forms.TextBox()
        Me.Label75 = New System.Windows.Forms.Label()
        Me.txtMaternoBen = New System.Windows.Forms.TextBox()
        Me.Label76 = New System.Windows.Forms.Label()
        Me.txtPaternoBen = New System.Windows.Forms.TextBox()
        Me.Label77 = New System.Windows.Forms.Label()
        Me.txtNombreBen = New System.Windows.Forms.TextBox()
        Me.Label79 = New System.Windows.Forms.Label()
        Me.CmdSiguiente = New System.Windows.Forms.Button()
        Me.nombreBen = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.paternoBen = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.maternoBen = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.rutBen = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dvBen = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.IdParentescoBen = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.nombreParentesco = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.porcentaje = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.fechaNacimientoBen = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dtBeneficiario, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.AliceBlue
        Me.GroupBox1.Controls.Add(Me.Label38)
        Me.GroupBox1.Controls.Add(Me.dtFechaNacBen)
        Me.GroupBox1.Controls.Add(Me.dtBeneficiario)
        Me.GroupBox1.Controls.Add(Me.Label70)
        Me.GroupBox1.Controls.Add(Me.txtPorcentajeBen)
        Me.GroupBox1.Controls.Add(Me.Label71)
        Me.GroupBox1.Controls.Add(Me.cmdEliminarBen)
        Me.GroupBox1.Controls.Add(Me.cmdModificarBen)
        Me.GroupBox1.Controls.Add(Me.cmdAgregarBen)
        Me.GroupBox1.Controls.Add(Me.Label72)
        Me.GroupBox1.Controls.Add(Me.Label73)
        Me.GroupBox1.Controls.Add(Me.cmbParentescoBen)
        Me.GroupBox1.Controls.Add(Me.txtDvBen)
        Me.GroupBox1.Controls.Add(Me.Label74)
        Me.GroupBox1.Controls.Add(Me.txtRutBen)
        Me.GroupBox1.Controls.Add(Me.Label75)
        Me.GroupBox1.Controls.Add(Me.txtMaternoBen)
        Me.GroupBox1.Controls.Add(Me.Label76)
        Me.GroupBox1.Controls.Add(Me.txtPaternoBen)
        Me.GroupBox1.Controls.Add(Me.Label77)
        Me.GroupBox1.Controls.Add(Me.txtNombreBen)
        Me.GroupBox1.Controls.Add(Me.Label79)
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.GroupBox1.Location = New System.Drawing.Point(18, 18)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(0)
        Me.GroupBox1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.GroupBox1.Size = New System.Drawing.Size(1438, 605)
        Me.GroupBox1.TabIndex = 53
        Me.GroupBox1.TabStop = False
        '
        'Label38
        '
        Me.Label38.BackColor = System.Drawing.Color.LightSeaGreen
        Me.Label38.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label38.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label38.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label38.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label38.Location = New System.Drawing.Point(381, 134)
        Me.Label38.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label38.Name = "Label38"
        Me.Label38.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label38.Size = New System.Drawing.Size(213, 38)
        Me.Label38.TabIndex = 288
        Me.Label38.Text = "Fecha Nacimiento"
        Me.Label38.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'dtFechaNacBen
        '
        Me.dtFechaNacBen.CustomFormat = ""
        Me.dtFechaNacBen.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtFechaNacBen.Location = New System.Drawing.Point(382, 177)
        Me.dtFechaNacBen.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.dtFechaNacBen.MinDate = New Date(1900, 1, 1, 0, 0, 0, 0)
        Me.dtFechaNacBen.Name = "dtFechaNacBen"
        Me.dtFechaNacBen.Size = New System.Drawing.Size(210, 26)
        Me.dtFechaNacBen.TabIndex = 287
        Me.dtFechaNacBen.Value = New Date(1900, 1, 1, 0, 0, 0, 0)
        '
        'dtBeneficiario
        '
        Me.dtBeneficiario.AllowUserToAddRows = False
        Me.dtBeneficiario.AllowUserToDeleteRows = False
        Me.dtBeneficiario.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dtBeneficiario.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.nombreBen, Me.paternoBen, Me.maternoBen, Me.rutBen, Me.dvBen, Me.IdParentescoBen, Me.nombreParentesco, Me.porcentaje, Me.fechaNacimientoBen})
        Me.dtBeneficiario.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.dtBeneficiario.Location = New System.Drawing.Point(9, 269)
        Me.dtBeneficiario.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.dtBeneficiario.Name = "dtBeneficiario"
        Me.dtBeneficiario.ReadOnly = True
        Me.dtBeneficiario.RowHeadersVisible = False
        Me.dtBeneficiario.RowHeadersWidth = 62
        Me.dtBeneficiario.Size = New System.Drawing.Size(1412, 312)
        Me.dtBeneficiario.TabIndex = 278
        '
        'Label70
        '
        Me.Label70.BackColor = System.Drawing.SystemColors.Control
        Me.Label70.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label70.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label70.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label70.Location = New System.Drawing.Point(330, 174)
        Me.Label70.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label70.Name = "Label70"
        Me.Label70.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label70.Size = New System.Drawing.Size(26, 26)
        Me.Label70.TabIndex = 274
        Me.Label70.Text = "%"
        '
        'txtPorcentajeBen
        '
        Me.txtPorcentajeBen.AcceptsReturn = True
        Me.txtPorcentajeBen.BackColor = System.Drawing.SystemColors.Window
        Me.txtPorcentajeBen.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPorcentajeBen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPorcentajeBen.Location = New System.Drawing.Point(250, 172)
        Me.txtPorcentajeBen.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtPorcentajeBen.MaxLength = 3
        Me.txtPorcentajeBen.Name = "txtPorcentajeBen"
        Me.txtPorcentajeBen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPorcentajeBen.Size = New System.Drawing.Size(66, 26)
        Me.txtPorcentajeBen.TabIndex = 273
        '
        'Label71
        '
        Me.Label71.BackColor = System.Drawing.SystemColors.Control
        Me.Label71.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label71.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label71.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label71.Location = New System.Drawing.Point(9, 22)
        Me.Label71.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label71.Name = "Label71"
        Me.Label71.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label71.Size = New System.Drawing.Size(261, 26)
        Me.Label71.TabIndex = 272
        Me.Label71.Text = "Ingrese Beneficiarios:"
        '
        'cmdEliminarBen
        '
        Me.cmdEliminarBen.Location = New System.Drawing.Point(1278, 223)
        Me.cmdEliminarBen.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmdEliminarBen.Name = "cmdEliminarBen"
        Me.cmdEliminarBen.Size = New System.Drawing.Size(147, 37)
        Me.cmdEliminarBen.TabIndex = 271
        Me.cmdEliminarBen.Text = "Eliminar"
        Me.cmdEliminarBen.UseVisualStyleBackColor = True
        '
        'cmdModificarBen
        '
        Me.cmdModificarBen.Location = New System.Drawing.Point(1276, 177)
        Me.cmdModificarBen.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmdModificarBen.Name = "cmdModificarBen"
        Me.cmdModificarBen.Size = New System.Drawing.Size(147, 37)
        Me.cmdModificarBen.TabIndex = 270
        Me.cmdModificarBen.Text = "Modificar"
        Me.cmdModificarBen.UseVisualStyleBackColor = True
        '
        'cmdAgregarBen
        '
        Me.cmdAgregarBen.Location = New System.Drawing.Point(1276, 132)
        Me.cmdAgregarBen.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmdAgregarBen.Name = "cmdAgregarBen"
        Me.cmdAgregarBen.Size = New System.Drawing.Size(147, 37)
        Me.cmdAgregarBen.TabIndex = 269
        Me.cmdAgregarBen.Text = "Agregar"
        Me.cmdAgregarBen.UseVisualStyleBackColor = True
        '
        'Label72
        '
        Me.Label72.BackColor = System.Drawing.Color.LightSeaGreen
        Me.Label72.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label72.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label72.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label72.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label72.Location = New System.Drawing.Point(250, 134)
        Me.Label72.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label72.Name = "Label72"
        Me.Label72.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label72.Size = New System.Drawing.Size(122, 38)
        Me.Label72.TabIndex = 268
        Me.Label72.Text = "Porcentaje"
        Me.Label72.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label73
        '
        Me.Label73.BackColor = System.Drawing.Color.LightSeaGreen
        Me.Label73.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label73.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label73.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label73.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label73.Location = New System.Drawing.Point(14, 131)
        Me.Label73.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label73.Name = "Label73"
        Me.Label73.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label73.Size = New System.Drawing.Size(224, 38)
        Me.Label73.TabIndex = 267
        Me.Label73.Text = "Parentesco"
        Me.Label73.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmbParentescoBen
        '
        Me.cmbParentescoBen.BackColor = System.Drawing.SystemColors.Window
        Me.cmbParentescoBen.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbParentescoBen.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbParentescoBen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbParentescoBen.Location = New System.Drawing.Point(14, 175)
        Me.cmbParentescoBen.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmbParentescoBen.Name = "cmbParentescoBen"
        Me.cmbParentescoBen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbParentescoBen.Size = New System.Drawing.Size(222, 28)
        Me.cmbParentescoBen.TabIndex = 266
        '
        'txtDvBen
        '
        Me.txtDvBen.AcceptsReturn = True
        Me.txtDvBen.BackColor = System.Drawing.SystemColors.Window
        Me.txtDvBen.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDvBen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDvBen.Location = New System.Drawing.Point(1362, 91)
        Me.txtDvBen.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtDvBen.MaxLength = 1
        Me.txtDvBen.Name = "txtDvBen"
        Me.txtDvBen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDvBen.Size = New System.Drawing.Size(56, 26)
        Me.txtDvBen.TabIndex = 265
        '
        'Label74
        '
        Me.Label74.BackColor = System.Drawing.Color.LightSeaGreen
        Me.Label74.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label74.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label74.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label74.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label74.Location = New System.Drawing.Point(1360, 48)
        Me.Label74.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label74.Name = "Label74"
        Me.Label74.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label74.Size = New System.Drawing.Size(60, 38)
        Me.Label74.TabIndex = 264
        Me.Label74.Text = "DV"
        Me.Label74.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtRutBen
        '
        Me.txtRutBen.AcceptsReturn = True
        Me.txtRutBen.BackColor = System.Drawing.SystemColors.Window
        Me.txtRutBen.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRutBen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRutBen.Location = New System.Drawing.Point(1221, 91)
        Me.txtRutBen.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtRutBen.MaxLength = 8
        Me.txtRutBen.Name = "txtRutBen"
        Me.txtRutBen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRutBen.Size = New System.Drawing.Size(133, 26)
        Me.txtRutBen.TabIndex = 263
        '
        'Label75
        '
        Me.Label75.BackColor = System.Drawing.Color.LightSeaGreen
        Me.Label75.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label75.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label75.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label75.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label75.Location = New System.Drawing.Point(1220, 48)
        Me.Label75.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label75.Name = "Label75"
        Me.Label75.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label75.Size = New System.Drawing.Size(136, 38)
        Me.Label75.TabIndex = 262
        Me.Label75.Text = "Rut"
        Me.Label75.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtMaternoBen
        '
        Me.txtMaternoBen.AcceptsReturn = True
        Me.txtMaternoBen.BackColor = System.Drawing.SystemColors.Window
        Me.txtMaternoBen.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtMaternoBen.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMaternoBen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMaternoBen.Location = New System.Drawing.Point(866, 91)
        Me.txtMaternoBen.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtMaternoBen.MaxLength = 50
        Me.txtMaternoBen.Name = "txtMaternoBen"
        Me.txtMaternoBen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMaternoBen.Size = New System.Drawing.Size(344, 26)
        Me.txtMaternoBen.TabIndex = 261
        '
        'Label76
        '
        Me.Label76.BackColor = System.Drawing.Color.LightSeaGreen
        Me.Label76.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label76.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label76.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label76.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label76.Location = New System.Drawing.Point(866, 48)
        Me.Label76.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label76.Name = "Label76"
        Me.Label76.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label76.Size = New System.Drawing.Size(345, 38)
        Me.Label76.TabIndex = 260
        Me.Label76.Text = "Materno"
        Me.Label76.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtPaternoBen
        '
        Me.txtPaternoBen.AcceptsReturn = True
        Me.txtPaternoBen.BackColor = System.Drawing.SystemColors.Window
        Me.txtPaternoBen.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPaternoBen.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPaternoBen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPaternoBen.Location = New System.Drawing.Point(525, 91)
        Me.txtPaternoBen.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtPaternoBen.MaxLength = 50
        Me.txtPaternoBen.Name = "txtPaternoBen"
        Me.txtPaternoBen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPaternoBen.Size = New System.Drawing.Size(330, 26)
        Me.txtPaternoBen.TabIndex = 259
        '
        'Label77
        '
        Me.Label77.BackColor = System.Drawing.Color.LightSeaGreen
        Me.Label77.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label77.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label77.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label77.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label77.Location = New System.Drawing.Point(525, 48)
        Me.Label77.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label77.Name = "Label77"
        Me.Label77.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label77.Size = New System.Drawing.Size(332, 38)
        Me.Label77.TabIndex = 258
        Me.Label77.Text = "Paterno"
        Me.Label77.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtNombreBen
        '
        Me.txtNombreBen.AcceptsReturn = True
        Me.txtNombreBen.BackColor = System.Drawing.SystemColors.Window
        Me.txtNombreBen.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtNombreBen.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNombreBen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNombreBen.Location = New System.Drawing.Point(15, 91)
        Me.txtNombreBen.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtNombreBen.MaxLength = 50
        Me.txtNombreBen.Name = "txtNombreBen"
        Me.txtNombreBen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNombreBen.Size = New System.Drawing.Size(476, 26)
        Me.txtNombreBen.TabIndex = 257
        '
        'Label79
        '
        Me.Label79.BackColor = System.Drawing.Color.LightSeaGreen
        Me.Label79.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label79.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label79.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label79.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label79.Location = New System.Drawing.Point(15, 48)
        Me.Label79.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label79.Name = "Label79"
        Me.Label79.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label79.Size = New System.Drawing.Size(482, 38)
        Me.Label79.TabIndex = 256
        Me.Label79.Text = "Nombres"
        Me.Label79.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'CmdSiguiente
        '
        Me.CmdSiguiente.BackColor = System.Drawing.SystemColors.Control
        Me.CmdSiguiente.Cursor = System.Windows.Forms.Cursors.Default
        Me.CmdSiguiente.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdSiguiente.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CmdSiguiente.Image = CType(resources.GetObject("CmdSiguiente.Image"), System.Drawing.Image)
        Me.CmdSiguiente.Location = New System.Drawing.Point(18, 632)
        Me.CmdSiguiente.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.CmdSiguiente.Name = "CmdSiguiente"
        Me.CmdSiguiente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CmdSiguiente.Size = New System.Drawing.Size(108, 97)
        Me.CmdSiguiente.TabIndex = 55
        Me.CmdSiguiente.Text = "&" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Guardar"
        Me.CmdSiguiente.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.CmdSiguiente.UseVisualStyleBackColor = False
        '
        'nombreBen
        '
        Me.nombreBen.HeaderText = "Nombre"
        Me.nombreBen.MinimumWidth = 8
        Me.nombreBen.Name = "nombreBen"
        Me.nombreBen.ReadOnly = True
        Me.nombreBen.Width = 150
        '
        'paternoBen
        '
        Me.paternoBen.HeaderText = "Paterno"
        Me.paternoBen.MinimumWidth = 8
        Me.paternoBen.Name = "paternoBen"
        Me.paternoBen.ReadOnly = True
        Me.paternoBen.Width = 150
        '
        'maternoBen
        '
        Me.maternoBen.HeaderText = "Materno"
        Me.maternoBen.MinimumWidth = 8
        Me.maternoBen.Name = "maternoBen"
        Me.maternoBen.ReadOnly = True
        Me.maternoBen.Width = 150
        '
        'rutBen
        '
        Me.rutBen.HeaderText = "Rut"
        Me.rutBen.MinimumWidth = 8
        Me.rutBen.Name = "rutBen"
        Me.rutBen.ReadOnly = True
        Me.rutBen.Width = 150
        '
        'dvBen
        '
        Me.dvBen.HeaderText = "DV"
        Me.dvBen.MinimumWidth = 8
        Me.dvBen.Name = "dvBen"
        Me.dvBen.ReadOnly = True
        Me.dvBen.Width = 150
        '
        'IdParentescoBen
        '
        Me.IdParentescoBen.HeaderText = "idParentesco"
        Me.IdParentescoBen.MinimumWidth = 8
        Me.IdParentescoBen.Name = "IdParentescoBen"
        Me.IdParentescoBen.ReadOnly = True
        Me.IdParentescoBen.Visible = False
        Me.IdParentescoBen.Width = 150
        '
        'nombreParentesco
        '
        Me.nombreParentesco.HeaderText = "Nombre Parentesco"
        Me.nombreParentesco.MinimumWidth = 8
        Me.nombreParentesco.Name = "nombreParentesco"
        Me.nombreParentesco.ReadOnly = True
        Me.nombreParentesco.Width = 150
        '
        'porcentaje
        '
        Me.porcentaje.HeaderText = "Porcentaje"
        Me.porcentaje.MinimumWidth = 8
        Me.porcentaje.Name = "porcentaje"
        Me.porcentaje.ReadOnly = True
        Me.porcentaje.Width = 150
        '
        'fechaNacimientoBen
        '
        Me.fechaNacimientoBen.HeaderText = "Fecha Nacimiento"
        Me.fechaNacimientoBen.MinimumWidth = 8
        Me.fechaNacimientoBen.Name = "fechaNacimientoBen"
        Me.fechaNacimientoBen.ReadOnly = True
        Me.fechaNacimientoBen.Width = 150
        '
        'frmBeneficiario
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.MediumSeaGreen
        Me.ClientSize = New System.Drawing.Size(1474, 748)
        Me.Controls.Add(Me.CmdSiguiente)
        Me.Controls.Add(Me.GroupBox1)
        Me.ForeColor = System.Drawing.Color.Black
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmBeneficiario"
        Me.Text = "frmBeneficiario"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dtBeneficiario, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Public WithEvents CmdSiguiente As System.Windows.Forms.Button
    Public WithEvents Label38 As System.Windows.Forms.Label
    Friend WithEvents dtFechaNacBen As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtBeneficiario As System.Windows.Forms.DataGridView
    Public WithEvents Label70 As System.Windows.Forms.Label
    Public WithEvents txtPorcentajeBen As System.Windows.Forms.TextBox
    Public WithEvents Label71 As System.Windows.Forms.Label
    Friend WithEvents cmdEliminarBen As System.Windows.Forms.Button
    Friend WithEvents cmdModificarBen As System.Windows.Forms.Button
    Friend WithEvents cmdAgregarBen As System.Windows.Forms.Button
    Public WithEvents Label72 As System.Windows.Forms.Label
    Public WithEvents Label73 As System.Windows.Forms.Label
    Public WithEvents cmbParentescoBen As System.Windows.Forms.ComboBox
    Public WithEvents txtDvBen As System.Windows.Forms.TextBox
    Public WithEvents Label74 As System.Windows.Forms.Label
    Public WithEvents txtRutBen As System.Windows.Forms.TextBox
    Public WithEvents Label75 As System.Windows.Forms.Label
    Public WithEvents txtMaternoBen As System.Windows.Forms.TextBox
    Public WithEvents Label76 As System.Windows.Forms.Label
    Public WithEvents txtPaternoBen As System.Windows.Forms.TextBox
    Public WithEvents Label77 As System.Windows.Forms.Label
    Public WithEvents txtNombreBen As System.Windows.Forms.TextBox
    Public WithEvents Label79 As System.Windows.Forms.Label
    Friend WithEvents nombreBen As DataGridViewTextBoxColumn
    Friend WithEvents paternoBen As DataGridViewTextBoxColumn
    Friend WithEvents maternoBen As DataGridViewTextBoxColumn
    Friend WithEvents rutBen As DataGridViewTextBoxColumn
    Friend WithEvents dvBen As DataGridViewTextBoxColumn
    Friend WithEvents IdParentescoBen As DataGridViewTextBoxColumn
    Friend WithEvents nombreParentesco As DataGridViewTextBoxColumn
    Friend WithEvents porcentaje As DataGridViewTextBoxColumn
    Friend WithEvents fechaNacimientoBen As DataGridViewTextBoxColumn
End Class
