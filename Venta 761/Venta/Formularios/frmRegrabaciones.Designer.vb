<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRegrabaciones
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
        Me.DataGridViewGes = New System.Windows.Forms.DataGridView()
        Me.btnSalirGes = New System.Windows.Forms.Button()
        Me.TextGesColor = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.id = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.rut = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dv = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.nombreCliente = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.fechaVenta = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.fechaEvaluacion = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.observacionesRechazo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.DataGridViewGes, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridViewGes
        '
        Me.DataGridViewGes.AllowUserToAddRows = False
        Me.DataGridViewGes.AllowUserToDeleteRows = False
        Me.DataGridViewGes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridViewGes.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.id, Me.rut, Me.dv, Me.nombreCliente, Me.fechaVenta, Me.fechaEvaluacion, Me.observacionesRechazo})
        Me.DataGridViewGes.Location = New System.Drawing.Point(12, 107)
        Me.DataGridViewGes.Name = "DataGridViewGes"
        Me.DataGridViewGes.ReadOnly = True
        Me.DataGridViewGes.Size = New System.Drawing.Size(638, 170)
        Me.DataGridViewGes.TabIndex = 0
        '
        'btnSalirGes
        '
        Me.btnSalirGes.Location = New System.Drawing.Point(544, 51)
        Me.btnSalirGes.Name = "btnSalirGes"
        Me.btnSalirGes.Size = New System.Drawing.Size(75, 23)
        Me.btnSalirGes.TabIndex = 1
        Me.btnSalirGes.Text = "Salir"
        Me.btnSalirGes.UseVisualStyleBackColor = True
        '
        'TextGesColor
        '
        Me.TextGesColor.Location = New System.Drawing.Point(152, 53)
        Me.TextGesColor.Name = "TextGesColor"
        Me.TextGesColor.Size = New System.Drawing.Size(73, 20)
        Me.TextGesColor.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(231, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(107, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Registros agendados"
        '
        'id
        '
        Me.id.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.id.DataPropertyName = "ID"
        Me.id.HeaderText = "ID"
        Me.id.Name = "id"
        Me.id.ReadOnly = True
        Me.id.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.id.Width = 43
        '
        'rut
        '
        Me.rut.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.rut.DataPropertyName = "RUT"
        Me.rut.HeaderText = "Rut"
        Me.rut.Name = "rut"
        Me.rut.ReadOnly = True
        Me.rut.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.rut.Width = 49
        '
        'dv
        '
        Me.dv.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.dv.DataPropertyName = "DV"
        Me.dv.HeaderText = "Dv"
        Me.dv.Name = "dv"
        Me.dv.ReadOnly = True
        Me.dv.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dv.Width = 46
        '
        'nombreCliente
        '
        Me.nombreCliente.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.nombreCliente.DataPropertyName = "NOMBRE"
        Me.nombreCliente.HeaderText = "Nombre Cliente"
        Me.nombreCliente.Name = "nombreCliente"
        Me.nombreCliente.ReadOnly = True
        Me.nombreCliente.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.nombreCliente.Width = 96
        '
        'fechaVenta
        '
        Me.fechaVenta.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.fechaVenta.DataPropertyName = "FECHA_VENTA"
        Me.fechaVenta.HeaderText = "Fecha Venta"
        Me.fechaVenta.Name = "fechaVenta"
        Me.fechaVenta.ReadOnly = True
        Me.fechaVenta.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.fechaVenta.Width = 86
        '
        'fechaEvaluacion
        '
        Me.fechaEvaluacion.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.fechaEvaluacion.DataPropertyName = "FECHA_EVALUACION"
        Me.fechaEvaluacion.HeaderText = "Fecha Evaluacion"
        Me.fechaEvaluacion.Name = "fechaEvaluacion"
        Me.fechaEvaluacion.ReadOnly = True
        Me.fechaEvaluacion.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.fechaEvaluacion.Width = 108
        '
        'observacionesRechazo
        '
        Me.observacionesRechazo.DataPropertyName = "OBSERVACIONES_RECHAZO"
        Me.observacionesRechazo.HeaderText = "Observaciones Rechazo"
        Me.observacionesRechazo.Name = "observacionesRechazo"
        Me.observacionesRechazo.ReadOnly = True
        Me.observacionesRechazo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.observacionesRechazo.Width = 250
        '
        'frmRegrabaciones
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(661, 292)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextGesColor)
        Me.Controls.Add(Me.btnSalirGes)
        Me.Controls.Add(Me.DataGridViewGes)
        Me.Name = "frmRegrabaciones"
        Me.Text = "Ventas Rechazadas en 1 Instancia"
        CType(Me.DataGridViewGes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DataGridViewGes As System.Windows.Forms.DataGridView
    Friend WithEvents btnSalirGes As System.Windows.Forms.Button
    Friend WithEvents TextGesColor As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents id As DataGridViewTextBoxColumn
    Friend WithEvents rut As DataGridViewTextBoxColumn
    Friend WithEvents dv As DataGridViewTextBoxColumn
    Friend WithEvents nombreCliente As DataGridViewTextBoxColumn
    Friend WithEvents fechaVenta As DataGridViewTextBoxColumn
    Friend WithEvents fechaEvaluacion As DataGridViewTextBoxColumn
    Friend WithEvents observacionesRechazo As DataGridViewTextBoxColumn
End Class
