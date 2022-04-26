<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEsperaDescanso
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEsperaDescanso))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblNombreDescanso = New System.Windows.Forms.Label()
        Me.btnFinDescanso = New System.Windows.Forms.Button()
        Me.pbImagenDescanso = New System.Windows.Forms.PictureBox()
        CType(Me.pbImagenDescanso, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.OrangeRed
        Me.Label1.Location = New System.Drawing.Point(112, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(512, 29)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "USTED SE ENCUENTRA EN TIEMPO DE: "
        '
        'lblNombreDescanso
        '
        Me.lblNombreDescanso.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNombreDescanso.ForeColor = System.Drawing.Color.Peru
        Me.lblNombreDescanso.Location = New System.Drawing.Point(112, 86)
        Me.lblNombreDescanso.Name = "lblNombreDescanso"
        Me.lblNombreDescanso.Size = New System.Drawing.Size(512, 29)
        Me.lblNombreDescanso.TabIndex = 1
        Me.lblNombreDescanso.Text = "CAPACITACION / RETROALIMENTACION"
        Me.lblNombreDescanso.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnFinDescanso
        '
        Me.btnFinDescanso.Font = New System.Drawing.Font("Comic Sans MS", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFinDescanso.ForeColor = System.Drawing.Color.DarkBlue
        Me.btnFinDescanso.Image = CType(resources.GetObject("btnFinDescanso.Image"), System.Drawing.Image)
        Me.btnFinDescanso.Location = New System.Drawing.Point(298, 340)
        Me.btnFinDescanso.Name = "btnFinDescanso"
        Me.btnFinDescanso.Size = New System.Drawing.Size(114, 74)
        Me.btnFinDescanso.TabIndex = 3
        Me.btnFinDescanso.Text = "Fin Descanso"
        Me.btnFinDescanso.TextImageRelation = System.Windows.Forms.TextImageRelation.TextAboveImage
        Me.btnFinDescanso.UseVisualStyleBackColor = True
        '
        'pbImagenDescanso
        '
        Me.pbImagenDescanso.Image = Global.APP_761.My.Resources.Resources.Colacion
        Me.pbImagenDescanso.ImageLocation = ""
        Me.pbImagenDescanso.Location = New System.Drawing.Point(298, 144)
        Me.pbImagenDescanso.Name = "pbImagenDescanso"
        Me.pbImagenDescanso.Size = New System.Drawing.Size(132, 141)
        Me.pbImagenDescanso.TabIndex = 2
        Me.pbImagenDescanso.TabStop = False
        '
        'frmEsperaDescanso
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.ClientSize = New System.Drawing.Size(747, 426)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnFinDescanso)
        Me.Controls.Add(Me.pbImagenDescanso)
        Me.Controls.Add(Me.lblNombreDescanso)
        Me.Controls.Add(Me.Label1)
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Name = "frmEsperaDescanso"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        CType(Me.pbImagenDescanso, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents lblNombreDescanso As Label
    Friend WithEvents pbImagenDescanso As PictureBox
    Friend WithEvents btnFinDescanso As Button
End Class
