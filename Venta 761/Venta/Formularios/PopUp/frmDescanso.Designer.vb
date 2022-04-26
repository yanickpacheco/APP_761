<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDescanso
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
        Me.rdbDescanso1 = New System.Windows.Forms.RadioButton()
        Me.rdbDescanso2 = New System.Windows.Forms.RadioButton()
        Me.rdbDescanso3 = New System.Windows.Forms.RadioButton()
        Me.rdbDescanso4 = New System.Windows.Forms.RadioButton()
        Me.btnActivarDescanso = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'rdbDescanso1
        '
        Me.rdbDescanso1.AutoSize = True
        Me.rdbDescanso1.ForeColor = System.Drawing.Color.Blue
        Me.rdbDescanso1.Location = New System.Drawing.Point(22, 19)
        Me.rdbDescanso1.Name = "rdbDescanso1"
        Me.rdbDescanso1.Size = New System.Drawing.Size(50, 17)
        Me.rdbDescanso1.TabIndex = 0
        Me.rdbDescanso1.TabStop = True
        Me.rdbDescanso1.Text = "Baño"
        Me.rdbDescanso1.UseVisualStyleBackColor = True
        '
        'rdbDescanso2
        '
        Me.rdbDescanso2.AutoSize = True
        Me.rdbDescanso2.ForeColor = System.Drawing.Color.Blue
        Me.rdbDescanso2.Location = New System.Drawing.Point(22, 52)
        Me.rdbDescanso2.Name = "rdbDescanso2"
        Me.rdbDescanso2.Size = New System.Drawing.Size(53, 17)
        Me.rdbDescanso2.TabIndex = 1
        Me.rdbDescanso2.TabStop = True
        Me.rdbDescanso2.Text = "Break"
        Me.rdbDescanso2.UseVisualStyleBackColor = True
        '
        'rdbDescanso3
        '
        Me.rdbDescanso3.AutoSize = True
        Me.rdbDescanso3.ForeColor = System.Drawing.Color.Blue
        Me.rdbDescanso3.Location = New System.Drawing.Point(22, 87)
        Me.rdbDescanso3.Name = "rdbDescanso3"
        Me.rdbDescanso3.Size = New System.Drawing.Size(66, 17)
        Me.rdbDescanso3.TabIndex = 2
        Me.rdbDescanso3.TabStop = True
        Me.rdbDescanso3.Text = "Colacion"
        Me.rdbDescanso3.UseVisualStyleBackColor = True
        '
        'rdbDescanso4
        '
        Me.rdbDescanso4.AutoSize = True
        Me.rdbDescanso4.ForeColor = System.Drawing.Color.Blue
        Me.rdbDescanso4.Location = New System.Drawing.Point(22, 124)
        Me.rdbDescanso4.Name = "rdbDescanso4"
        Me.rdbDescanso4.Size = New System.Drawing.Size(178, 17)
        Me.rdbDescanso4.TabIndex = 3
        Me.rdbDescanso4.TabStop = True
        Me.rdbDescanso4.Text = "RetroAlimentacion/Capacitacion"
        Me.rdbDescanso4.UseVisualStyleBackColor = True
        '
        'btnActivarDescanso
        '
        Me.btnActivarDescanso.Font = New System.Drawing.Font("Comic Sans MS", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnActivarDescanso.Location = New System.Drawing.Point(98, 204)
        Me.btnActivarDescanso.Name = "btnActivarDescanso"
        Me.btnActivarDescanso.Size = New System.Drawing.Size(97, 48)
        Me.btnActivarDescanso.TabIndex = 4
        Me.btnActivarDescanso.Text = "Activar Descanso"
        Me.btnActivarDescanso.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.MidnightBlue
        Me.Label1.Location = New System.Drawing.Point(58, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(195, 17)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Escoga Tiempo Descanso"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rdbDescanso1)
        Me.GroupBox1.Controls.Add(Me.rdbDescanso2)
        Me.GroupBox1.Controls.Add(Me.rdbDescanso3)
        Me.GroupBox1.Controls.Add(Me.rdbDescanso4)
        Me.GroupBox1.Location = New System.Drawing.Point(39, 35)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(217, 163)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        '
        'btnSalir
        '
        Me.btnSalir.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSalir.Location = New System.Drawing.Point(218, 259)
        Me.btnSalir.Name = "btnSalir"
        Me.btnSalir.Size = New System.Drawing.Size(66, 24)
        Me.btnSalir.TabIndex = 7
        Me.btnSalir.Text = "Salir"
        Me.btnSalir.UseVisualStyleBackColor = True
        '
        'frmDescanso
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(294, 291)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnSalir)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnActivarDescanso)
        Me.Name = "frmDescanso"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Descanso"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents rdbDescanso1 As RadioButton
    Friend WithEvents rdbDescanso2 As RadioButton
    Friend WithEvents rdbDescanso3 As RadioButton
    Friend WithEvents rdbDescanso4 As RadioButton
    Friend WithEvents btnActivarDescanso As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents btnSalir As Button
End Class
