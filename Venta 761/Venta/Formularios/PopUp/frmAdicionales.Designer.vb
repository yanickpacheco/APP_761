<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAdicionales
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAdicionales))
        Me.lblPesosUF = New System.Windows.Forms.Label()
        Me.lblTotalPrimaUf = New System.Windows.Forms.Label()
        Me.lblPrevisionAdic = New System.Windows.Forms.Label()
        Me.cmbSaludAdic = New System.Windows.Forms.ComboBox()
        Me.dtAdicional = New System.Windows.Forms.DataGridView()
        Me.lblSexoA = New System.Windows.Forms.Label()
        Me.cmbSexoA = New System.Windows.Forms.ComboBox()
        Me.dtFechaNacAdic = New System.Windows.Forms.DateTimePicker()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.btnModificar = New System.Windows.Forms.Button()
        Me.btnAgregar = New System.Windows.Forms.Button()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.Parentesco = New System.Windows.Forms.Label()
        Me.cmbParentescoAdic = New System.Windows.Forms.ComboBox()
        Me.txtDvA = New System.Windows.Forms.TextBox()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.txtRutA = New System.Windows.Forms.TextBox()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.txtMaternoA = New System.Windows.Forms.TextBox()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.txtPaternoA = New System.Windows.Forms.TextBox()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.txtNombreA = New System.Windows.Forms.TextBox()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.CmdSiguiente = New System.Windows.Forms.Button()
        Me.nombre = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.paterno = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.materno = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.rut = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dv = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Sexo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.tipo_parentesco = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.fechaNacimiento = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.idParentescoAdi = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.valorUf = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.valorPesos = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.idPlan = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SistemaSalud = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.dtAdicional, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblPesosUF
        '
        Me.lblPesosUF.AutoSize = True
        Me.lblPesosUF.Location = New System.Drawing.Point(713, 138)
        Me.lblPesosUF.Name = "lblPesosUF"
        Me.lblPesosUF.Size = New System.Drawing.Size(51, 17)
        Me.lblPesosUF.TabIndex = 248
        Me.lblPesosUF.Text = "Label1"
        '
        'lblTotalPrimaUf
        '
        Me.lblTotalPrimaUf.AutoSize = True
        Me.lblTotalPrimaUf.Location = New System.Drawing.Point(712, 163)
        Me.lblTotalPrimaUf.Name = "lblTotalPrimaUf"
        Me.lblTotalPrimaUf.Size = New System.Drawing.Size(51, 17)
        Me.lblTotalPrimaUf.TabIndex = 247
        Me.lblTotalPrimaUf.Text = "Label1"
        '
        'lblPrevisionAdic
        '
        Me.lblPrevisionAdic.BackColor = System.Drawing.Color.LightSeaGreen
        Me.lblPrevisionAdic.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblPrevisionAdic.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPrevisionAdic.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPrevisionAdic.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblPrevisionAdic.Location = New System.Drawing.Point(456, 129)
        Me.lblPrevisionAdic.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblPrevisionAdic.Name = "lblPrevisionAdic"
        Me.lblPrevisionAdic.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPrevisionAdic.Size = New System.Drawing.Size(233, 31)
        Me.lblPrevisionAdic.TabIndex = 246
        Me.lblPrevisionAdic.Text = "Prevision"
        Me.lblPrevisionAdic.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmbSaludAdic
        '
        Me.cmbSaludAdic.BackColor = System.Drawing.SystemColors.Window
        Me.cmbSaludAdic.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbSaludAdic.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSaludAdic.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbSaludAdic.Items.AddRange(New Object() {"[No Especificado]", "BANMEDICA S.A.", "CAPREDENA", "CHUQUICAMATA LTDA.", "COLMENA GOLDEN CROSS S.A.", "CONSALUD S.A.", "CRUZ BLANCA S.A.", "DIPRECA", "FUNDACION LTDA.", "FUSAT LTDA.", "NUEVA MASVIDA S.A.", "RIO BLANCO LTDA.", "SAN LORENZO LTDA.", "VIDA TRES S.A.", "FONASA TRAMO B", "FONASA TRAMO C", "FONASA TRAMO D"})
        Me.cmbSaludAdic.Location = New System.Drawing.Point(456, 163)
        Me.cmbSaludAdic.Margin = New System.Windows.Forms.Padding(4)
        Me.cmbSaludAdic.Name = "cmbSaludAdic"
        Me.cmbSaludAdic.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbSaludAdic.Size = New System.Drawing.Size(233, 24)
        Me.cmbSaludAdic.TabIndex = 245
        '
        'dtAdicional
        '
        Me.dtAdicional.AllowUserToAddRows = False
        Me.dtAdicional.AllowUserToDeleteRows = False
        Me.dtAdicional.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dtAdicional.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.nombre, Me.paterno, Me.materno, Me.rut, Me.dv, Me.Sexo, Me.tipo_parentesco, Me.fechaNacimiento, Me.idParentescoAdi, Me.valorUf, Me.valorPesos, Me.idPlan, Me.SistemaSalud})
        Me.dtAdicional.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.dtAdicional.Location = New System.Drawing.Point(19, 221)
        Me.dtAdicional.Margin = New System.Windows.Forms.Padding(4)
        Me.dtAdicional.Name = "dtAdicional"
        Me.dtAdicional.ReadOnly = True
        Me.dtAdicional.RowHeadersVisible = False
        Me.dtAdicional.RowHeadersWidth = 51
        Me.dtAdicional.Size = New System.Drawing.Size(1204, 234)
        Me.dtAdicional.TabIndex = 244
        '
        'lblSexoA
        '
        Me.lblSexoA.BackColor = System.Drawing.Color.LightSeaGreen
        Me.lblSexoA.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSexoA.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSexoA.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSexoA.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblSexoA.Location = New System.Drawing.Point(1102, 53)
        Me.lblSexoA.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblSexoA.Name = "lblSexoA"
        Me.lblSexoA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSexoA.Size = New System.Drawing.Size(147, 31)
        Me.lblSexoA.TabIndex = 243
        Me.lblSexoA.Text = "Sexo"
        Me.lblSexoA.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmbSexoA
        '
        Me.cmbSexoA.BackColor = System.Drawing.SystemColors.Window
        Me.cmbSexoA.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbSexoA.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSexoA.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbSexoA.Items.AddRange(New Object() {"--Seleccione--", "M", "F"})
        Me.cmbSexoA.Location = New System.Drawing.Point(1102, 88)
        Me.cmbSexoA.Margin = New System.Windows.Forms.Padding(4)
        Me.cmbSexoA.Name = "cmbSexoA"
        Me.cmbSexoA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbSexoA.Size = New System.Drawing.Size(145, 24)
        Me.cmbSexoA.TabIndex = 231
        '
        'dtFechaNacAdic
        '
        Me.dtFechaNacAdic.CustomFormat = ""
        Me.dtFechaNacAdic.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtFechaNacAdic.Location = New System.Drawing.Point(19, 164)
        Me.dtFechaNacAdic.Margin = New System.Windows.Forms.Padding(4)
        Me.dtFechaNacAdic.MinDate = New Date(1900, 1, 1, 0, 0, 0, 0)
        Me.dtFechaNacAdic.Name = "dtFechaNacAdic"
        Me.dtFechaNacAdic.Size = New System.Drawing.Size(187, 22)
        Me.dtFechaNacAdic.TabIndex = 233
        Me.dtFechaNacAdic.Value = New Date(2014, 9, 4, 0, 0, 0, 0)
        '
        'Label31
        '
        Me.Label31.BackColor = System.Drawing.Color.AliceBlue
        Me.Label31.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label31.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label31.ForeColor = System.Drawing.Color.FromArgb(CType(CType(100, Byte), Integer), CType(CType(15, Byte), Integer), CType(CType(91, Byte), Integer))
        Me.Label31.Location = New System.Drawing.Point(29, 18)
        Me.Label31.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label31.Name = "Label31"
        Me.Label31.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label31.Size = New System.Drawing.Size(196, 21)
        Me.Label31.TabIndex = 242
        Me.Label31.Text = "Ingrese Adicionales:"
        '
        'btnEliminar
        '
        Me.btnEliminar.Location = New System.Drawing.Point(1117, 164)
        Me.btnEliminar.Margin = New System.Windows.Forms.Padding(4)
        Me.btnEliminar.Name = "btnEliminar"
        Me.btnEliminar.Size = New System.Drawing.Size(131, 30)
        Me.btnEliminar.TabIndex = 241
        Me.btnEliminar.Text = "Eliminar"
        Me.btnEliminar.UseVisualStyleBackColor = True
        '
        'btnModificar
        '
        Me.btnModificar.Location = New System.Drawing.Point(978, 163)
        Me.btnModificar.Margin = New System.Windows.Forms.Padding(4)
        Me.btnModificar.Name = "btnModificar"
        Me.btnModificar.Size = New System.Drawing.Size(131, 30)
        Me.btnModificar.TabIndex = 240
        Me.btnModificar.Text = "Modificar"
        Me.btnModificar.UseVisualStyleBackColor = True
        '
        'btnAgregar
        '
        Me.btnAgregar.Location = New System.Drawing.Point(839, 163)
        Me.btnAgregar.Margin = New System.Windows.Forms.Padding(4)
        Me.btnAgregar.Name = "btnAgregar"
        Me.btnAgregar.Size = New System.Drawing.Size(131, 30)
        Me.btnAgregar.TabIndex = 236
        Me.btnAgregar.Text = "Agregar"
        Me.btnAgregar.UseVisualStyleBackColor = True
        '
        'Label32
        '
        Me.Label32.BackColor = System.Drawing.Color.LightSeaGreen
        Me.Label32.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label32.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label32.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label32.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label32.Location = New System.Drawing.Point(19, 129)
        Me.Label32.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label32.Name = "Label32"
        Me.Label32.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label32.Size = New System.Drawing.Size(191, 31)
        Me.Label32.TabIndex = 239
        Me.Label32.Text = "Fecha de Nacimiento"
        Me.Label32.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Parentesco
        '
        Me.Parentesco.BackColor = System.Drawing.Color.LightSeaGreen
        Me.Parentesco.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Parentesco.Cursor = System.Windows.Forms.Cursors.Default
        Me.Parentesco.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Parentesco.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Parentesco.Location = New System.Drawing.Point(218, 129)
        Me.Parentesco.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Parentesco.Name = "Parentesco"
        Me.Parentesco.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Parentesco.Size = New System.Drawing.Size(223, 31)
        Me.Parentesco.TabIndex = 238
        Me.Parentesco.Text = "Parentesco"
        Me.Parentesco.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmbParentescoAdic
        '
        Me.cmbParentescoAdic.BackColor = System.Drawing.SystemColors.Window
        Me.cmbParentescoAdic.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbParentescoAdic.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbParentescoAdic.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbParentescoAdic.Location = New System.Drawing.Point(218, 164)
        Me.cmbParentescoAdic.Margin = New System.Windows.Forms.Padding(4)
        Me.cmbParentescoAdic.Name = "cmbParentescoAdic"
        Me.cmbParentescoAdic.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbParentescoAdic.Size = New System.Drawing.Size(223, 24)
        Me.cmbParentescoAdic.TabIndex = 234
        '
        'txtDvA
        '
        Me.txtDvA.AcceptsReturn = True
        Me.txtDvA.BackColor = System.Drawing.SystemColors.Window
        Me.txtDvA.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDvA.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDvA.Location = New System.Drawing.Point(1042, 88)
        Me.txtDvA.Margin = New System.Windows.Forms.Padding(4)
        Me.txtDvA.MaxLength = 1
        Me.txtDvA.Name = "txtDvA"
        Me.txtDvA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDvA.Size = New System.Drawing.Size(51, 22)
        Me.txtDvA.TabIndex = 230
        '
        'Label29
        '
        Me.Label29.BackColor = System.Drawing.Color.LightSeaGreen
        Me.Label29.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label29.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label29.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label29.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label29.Location = New System.Drawing.Point(1040, 53)
        Me.Label29.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label29.Name = "Label29"
        Me.Label29.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label29.Size = New System.Drawing.Size(53, 31)
        Me.Label29.TabIndex = 237
        Me.Label29.Text = "DV"
        Me.Label29.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtRutA
        '
        Me.txtRutA.AcceptsReturn = True
        Me.txtRutA.BackColor = System.Drawing.SystemColors.Window
        Me.txtRutA.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRutA.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRutA.Location = New System.Drawing.Point(916, 88)
        Me.txtRutA.Margin = New System.Windows.Forms.Padding(4)
        Me.txtRutA.MaxLength = 8
        Me.txtRutA.Name = "txtRutA"
        Me.txtRutA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRutA.Size = New System.Drawing.Size(119, 22)
        Me.txtRutA.TabIndex = 229
        '
        'Label30
        '
        Me.Label30.BackColor = System.Drawing.Color.LightSeaGreen
        Me.Label30.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label30.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label30.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label30.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label30.Location = New System.Drawing.Point(915, 53)
        Me.Label30.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label30.Name = "Label30"
        Me.Label30.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label30.Size = New System.Drawing.Size(121, 31)
        Me.Label30.TabIndex = 235
        Me.Label30.Text = "Rut"
        Me.Label30.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtMaternoA
        '
        Me.txtMaternoA.AcceptsReturn = True
        Me.txtMaternoA.BackColor = System.Drawing.SystemColors.Window
        Me.txtMaternoA.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtMaternoA.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMaternoA.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMaternoA.Location = New System.Drawing.Point(698, 88)
        Me.txtMaternoA.Margin = New System.Windows.Forms.Padding(4)
        Me.txtMaternoA.MaxLength = 50
        Me.txtMaternoA.Name = "txtMaternoA"
        Me.txtMaternoA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMaternoA.Size = New System.Drawing.Size(209, 22)
        Me.txtMaternoA.TabIndex = 227
        '
        'Label28
        '
        Me.Label28.BackColor = System.Drawing.Color.LightSeaGreen
        Me.Label28.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label28.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label28.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label28.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label28.Location = New System.Drawing.Point(698, 53)
        Me.Label28.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label28.Name = "Label28"
        Me.Label28.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label28.Size = New System.Drawing.Size(209, 31)
        Me.Label28.TabIndex = 232
        Me.Label28.Text = "Materno"
        Me.Label28.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtPaternoA
        '
        Me.txtPaternoA.AcceptsReturn = True
        Me.txtPaternoA.BackColor = System.Drawing.SystemColors.Window
        Me.txtPaternoA.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPaternoA.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPaternoA.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPaternoA.Location = New System.Drawing.Point(471, 88)
        Me.txtPaternoA.Margin = New System.Windows.Forms.Padding(4)
        Me.txtPaternoA.MaxLength = 50
        Me.txtPaternoA.Name = "txtPaternoA"
        Me.txtPaternoA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPaternoA.Size = New System.Drawing.Size(217, 22)
        Me.txtPaternoA.TabIndex = 226
        '
        'Label27
        '
        Me.Label27.BackColor = System.Drawing.Color.LightSeaGreen
        Me.Label27.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label27.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label27.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label27.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label27.Location = New System.Drawing.Point(471, 53)
        Me.Label27.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label27.Name = "Label27"
        Me.Label27.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label27.Size = New System.Drawing.Size(219, 31)
        Me.Label27.TabIndex = 228
        Me.Label27.Text = "Paterno"
        Me.Label27.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtNombreA
        '
        Me.txtNombreA.AcceptsReturn = True
        Me.txtNombreA.BackColor = System.Drawing.SystemColors.Window
        Me.txtNombreA.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtNombreA.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNombreA.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNombreA.Location = New System.Drawing.Point(19, 88)
        Me.txtNombreA.Margin = New System.Windows.Forms.Padding(4)
        Me.txtNombreA.MaxLength = 50
        Me.txtNombreA.Name = "txtNombreA"
        Me.txtNombreA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNombreA.Size = New System.Drawing.Size(440, 22)
        Me.txtNombreA.TabIndex = 225
        '
        'Label25
        '
        Me.Label25.BackColor = System.Drawing.Color.LightSeaGreen
        Me.Label25.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label25.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label25.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label25.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label25.Location = New System.Drawing.Point(19, 53)
        Me.Label25.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label25.Name = "Label25"
        Me.Label25.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label25.Size = New System.Drawing.Size(444, 31)
        Me.Label25.TabIndex = 224
        Me.Label25.Text = "Nombre 1"
        Me.Label25.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.AliceBlue
        Me.GroupBox1.Controls.Add(Me.Label31)
        Me.GroupBox1.Controls.Add(Me.dtAdicional)
        Me.GroupBox1.Controls.Add(Me.lblPesosUF)
        Me.GroupBox1.Controls.Add(Me.lblSexoA)
        Me.GroupBox1.Controls.Add(Me.lblTotalPrimaUf)
        Me.GroupBox1.Controls.Add(Me.Label25)
        Me.GroupBox1.Controls.Add(Me.lblPrevisionAdic)
        Me.GroupBox1.Controls.Add(Me.txtNombreA)
        Me.GroupBox1.Controls.Add(Me.cmbSaludAdic)
        Me.GroupBox1.Controls.Add(Me.Label27)
        Me.GroupBox1.Controls.Add(Me.txtPaternoA)
        Me.GroupBox1.Controls.Add(Me.dtFechaNacAdic)
        Me.GroupBox1.Controls.Add(Me.Label28)
        Me.GroupBox1.Controls.Add(Me.btnEliminar)
        Me.GroupBox1.Controls.Add(Me.cmbSexoA)
        Me.GroupBox1.Controls.Add(Me.btnModificar)
        Me.GroupBox1.Controls.Add(Me.txtMaternoA)
        Me.GroupBox1.Controls.Add(Me.btnAgregar)
        Me.GroupBox1.Controls.Add(Me.Label30)
        Me.GroupBox1.Controls.Add(Me.Label32)
        Me.GroupBox1.Controls.Add(Me.Parentesco)
        Me.GroupBox1.Controls.Add(Me.txtRutA)
        Me.GroupBox1.Controls.Add(Me.cmbParentescoAdic)
        Me.GroupBox1.Controls.Add(Me.Label29)
        Me.GroupBox1.Controls.Add(Me.txtDvA)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1277, 471)
        Me.GroupBox1.TabIndex = 249
        Me.GroupBox1.TabStop = False
        '
        'CmdSiguiente
        '
        Me.CmdSiguiente.BackColor = System.Drawing.SystemColors.Control
        Me.CmdSiguiente.Cursor = System.Windows.Forms.Cursors.Default
        Me.CmdSiguiente.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdSiguiente.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CmdSiguiente.Image = CType(resources.GetObject("CmdSiguiente.Image"), System.Drawing.Image)
        Me.CmdSiguiente.Location = New System.Drawing.Point(13, 490)
        Me.CmdSiguiente.Margin = New System.Windows.Forms.Padding(4)
        Me.CmdSiguiente.Name = "CmdSiguiente"
        Me.CmdSiguiente.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CmdSiguiente.Size = New System.Drawing.Size(96, 78)
        Me.CmdSiguiente.TabIndex = 250
        Me.CmdSiguiente.Text = "&" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Guardar"
        Me.CmdSiguiente.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.CmdSiguiente.UseVisualStyleBackColor = False
        '
        'nombre
        '
        Me.nombre.HeaderText = "Nombre"
        Me.nombre.MinimumWidth = 6
        Me.nombre.Name = "nombre"
        Me.nombre.ReadOnly = True
        Me.nombre.Width = 125
        '
        'paterno
        '
        Me.paterno.HeaderText = "Paterno"
        Me.paterno.MinimumWidth = 6
        Me.paterno.Name = "paterno"
        Me.paterno.ReadOnly = True
        Me.paterno.Width = 125
        '
        'materno
        '
        Me.materno.HeaderText = "Materno"
        Me.materno.MinimumWidth = 6
        Me.materno.Name = "materno"
        Me.materno.ReadOnly = True
        Me.materno.Width = 125
        '
        'rut
        '
        Me.rut.HeaderText = "Rut"
        Me.rut.MinimumWidth = 6
        Me.rut.Name = "rut"
        Me.rut.ReadOnly = True
        Me.rut.Width = 125
        '
        'dv
        '
        Me.dv.HeaderText = "DV"
        Me.dv.MinimumWidth = 6
        Me.dv.Name = "dv"
        Me.dv.ReadOnly = True
        Me.dv.Width = 125
        '
        'Sexo
        '
        Me.Sexo.HeaderText = "Sexo"
        Me.Sexo.MinimumWidth = 6
        Me.Sexo.Name = "Sexo"
        Me.Sexo.ReadOnly = True
        Me.Sexo.Width = 125
        '
        'tipo_parentesco
        '
        Me.tipo_parentesco.HeaderText = "Parentesco"
        Me.tipo_parentesco.MinimumWidth = 6
        Me.tipo_parentesco.Name = "tipo_parentesco"
        Me.tipo_parentesco.ReadOnly = True
        Me.tipo_parentesco.Width = 125
        '
        'fechaNacimiento
        '
        Me.fechaNacimiento.HeaderText = "Fecha Nacimiento"
        Me.fechaNacimiento.MinimumWidth = 6
        Me.fechaNacimiento.Name = "fechaNacimiento"
        Me.fechaNacimiento.ReadOnly = True
        Me.fechaNacimiento.Width = 125
        '
        'idParentescoAdi
        '
        Me.idParentescoAdi.HeaderText = "idParentesco"
        Me.idParentescoAdi.MinimumWidth = 6
        Me.idParentescoAdi.Name = "idParentescoAdi"
        Me.idParentescoAdi.ReadOnly = True
        Me.idParentescoAdi.Visible = False
        Me.idParentescoAdi.Width = 125
        '
        'valorUf
        '
        Me.valorUf.HeaderText = "valorUf"
        Me.valorUf.MinimumWidth = 6
        Me.valorUf.Name = "valorUf"
        Me.valorUf.ReadOnly = True
        Me.valorUf.Width = 125
        '
        'valorPesos
        '
        Me.valorPesos.HeaderText = "valorPesos"
        Me.valorPesos.MinimumWidth = 6
        Me.valorPesos.Name = "valorPesos"
        Me.valorPesos.ReadOnly = True
        Me.valorPesos.Width = 125
        '
        'idPlan
        '
        Me.idPlan.HeaderText = "idPlan"
        Me.idPlan.MinimumWidth = 6
        Me.idPlan.Name = "idPlan"
        Me.idPlan.ReadOnly = True
        Me.idPlan.Width = 125
        '
        'SistemaSalud
        '
        Me.SistemaSalud.HeaderText = "SistemaSalud"
        Me.SistemaSalud.MinimumWidth = 6
        Me.SistemaSalud.Name = "SistemaSalud"
        Me.SistemaSalud.ReadOnly = True
        Me.SistemaSalud.Width = 125
        '
        'frmAdicionales
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.BackColor = System.Drawing.Color.MediumSeaGreen
        Me.ClientSize = New System.Drawing.Size(1327, 571)
        Me.Controls.Add(Me.CmdSiguiente)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "frmAdicionales"
        Me.Text = "frmAdicionales"
        CType(Me.dtAdicional, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblPesosUF As Label
    Friend WithEvents lblTotalPrimaUf As Label
    Public WithEvents lblPrevisionAdic As Label
    Public WithEvents cmbSaludAdic As ComboBox
    Friend WithEvents dtAdicional As DataGridView
    Public WithEvents lblSexoA As Label
    Public WithEvents cmbSexoA As ComboBox
    Friend WithEvents dtFechaNacAdic As DateTimePicker
    Public WithEvents Label31 As Label
    Friend WithEvents btnEliminar As Button
    Friend WithEvents btnModificar As Button
    Friend WithEvents btnAgregar As Button
    Public WithEvents Label32 As Label
    Public WithEvents Parentesco As Label
    Public WithEvents cmbParentescoAdic As ComboBox
    Public WithEvents txtDvA As TextBox
    Public WithEvents Label29 As Label
    Public WithEvents txtRutA As TextBox
    Public WithEvents Label30 As Label
    Public WithEvents txtMaternoA As TextBox
    Public WithEvents Label28 As Label
    Public WithEvents txtPaternoA As TextBox
    Public WithEvents Label27 As Label
    Public WithEvents txtNombreA As TextBox
    Public WithEvents Label25 As Label
    Friend WithEvents GroupBox1 As GroupBox
    Public WithEvents CmdSiguiente As Button
    Friend WithEvents nombre As DataGridViewTextBoxColumn
    Friend WithEvents paterno As DataGridViewTextBoxColumn
    Friend WithEvents materno As DataGridViewTextBoxColumn
    Friend WithEvents rut As DataGridViewTextBoxColumn
    Friend WithEvents dv As DataGridViewTextBoxColumn
    Friend WithEvents Sexo As DataGridViewTextBoxColumn
    Friend WithEvents tipo_parentesco As DataGridViewTextBoxColumn
    Friend WithEvents fechaNacimiento As DataGridViewTextBoxColumn
    Friend WithEvents idParentescoAdi As DataGridViewTextBoxColumn
    Friend WithEvents valorUf As DataGridViewTextBoxColumn
    Friend WithEvents valorPesos As DataGridViewTextBoxColumn
    Friend WithEvents idPlan As DataGridViewTextBoxColumn
    Friend WithEvents SistemaSalud As DataGridViewTextBoxColumn
End Class
