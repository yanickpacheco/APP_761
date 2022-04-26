Imports Entidad
Imports BI

Public Class frmRegrabaciones
    Inherits System.Windows.Forms.Form
    Dim opc As Short
    Dim biRes As New clsRegrabacionesBI
    Private Sub frmRegrabaciones_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'frmLogin.Close()
        frmVenta.Hide() 'ocultamos el formulario principal
        opc = 0
        Dim table As New DataTable
        table = biRes.Regrabaciones(WS_IDUSUARIO)
        DataGridViewGes.DataSource = table
        'TextGesColor.BackColor = Color.Aqua
    End Sub

    '***********Al hacer doble click obtenemos la id del cliente a evaluar para ir a evaluar*********************
    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewGes.CellDoubleClick
        'ModGeneral.conectarTelefonia()
        'Logear_Usuario(WS_IDUSUARIO, 1)
        GesId = DataGridViewGes.Rows(e.RowIndex).Cells("ID").Value
        claveRegistroActual = GesId
        estado_perfil = True
        opc = 1
        frmVenta.Show()
        Me.Close()
    End Sub
    '****************Salir de la aplicacion******************
    Private Sub btnSalirGes_Click(sender As Object, e As EventArgs) Handles btnSalirGes.Click
        If MsgBox("¿Está seguro que desea salir de la aplicación?", MsgBoxStyle.YesNo, "CORDIALPHONE") = MsgBoxResult.Yes Then
            Logear_Usuario(WS_IDUSUARIO, 2)
            'If db_central = 4 Then
            '    vpPosicion.LogoutTelefonia((vpPosicion.Usuario))
            'End If
            End
        End If
    End Sub

    Private Sub frmRegrabaciones_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        If opc = 0 Then
            If MsgBox("¿Está seguro que desea salir de la aplicación?", MsgBoxStyle.YesNo, "CORDIALPHONE") = MsgBoxResult.Yes Then
                Logear_Usuario(WS_IDUSUARIO, 2)
                If db_central = 4 Then
                    vpPosicion.LogoutTelefonia((vpPosicion.Usuario))
                End If
                End
            End If
        End If
    End Sub


End Class