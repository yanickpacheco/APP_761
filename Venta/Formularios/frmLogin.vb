Option Strict Off
Option Explicit On
Imports BI

Friend Class frmLogin
    Inherits System.Windows.Forms.Form
    Dim opc As Short
    Public tablaUsuarios As New DataTable

    '*******metodo al presionar boton aceptar de boton login****************
    Private Sub btnAceptar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnAceptar.Click
        If cmbUsuarios.SelectedIndex < 1 Or cmbUsuarios.Text = "" Then
            MsgBox("Debe seleccionar un usuario", MsgBoxStyle.Exclamation, "Atención")
            Exit Sub
        End If
        opc = 1
        Dim valido As Boolean
        Dim biUsuario As New clsUsuarioBI
        WS_IDUSUARIO = cmbUsuarios.Text
        If CheckBoxLogin.Checked Then ' si checkbox esta marcado es para entrar como regrabador
            valido = biUsuario.Validar_Reg(WS_IDUSUARIO) ' validamos si el usuario tiene  el perfil de sistemas
            If valido Then
                perfil = "Regrabador"
            End If
            If Not valido Then
                MsgBox("El usuario no tiene permisos como Ejecutivo de Regrabación", MsgBoxStyle.Exclamation, "Atención")
                Exit Sub
            End If
        ElseIf Not CheckBoxLogin.Checked Then ' si checkbox no esta marcado entra como ventas
            valido = biUsuario.Validar_user(WS_IDUSUARIO) ' validamos el usuario para ventas
            If valido Then
                perfil = "Ejecutivo"
            Else
                MsgBox("El usuario no tiene permisos como Ejecutivo de Ventas", MsgBoxStyle.Exclamation, "Atención")
                Exit Sub
            End If
        End If
        If WS_IDUSUARIO <> "ADRIAN.LEBLANC" Then
            If ConectaNeotel((txtusuxlite.Text)) Then
                ModGeneral.conectarTelefonia()
                Logear_Usuario(WS_IDUSUARIO, 1)
                frmVenta.Show()
                Me.Close()
            End If
        Else
            frmVenta.Show()
            Me.Close()
        End If
    End Sub

    Private Sub frmLogin_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Me.Text = "NEW Versión: " & My.Application.Info.Version.Major.ToString & "." & My.Application.Info.Version.Minor.ToString _
      & "." & My.Application.Info.Version.Revision.ToString & "-" & csNombreAplicacion
        Dim vgFunComunes As New clsFuncionesComunes
        Dim us As New clsListaUsuarioBaseBI
        tablaUsuarios = us.ListarUsuarios(vgCampania.idCRM)

        vgFunComunes.llenaComboboxTabla(cmbUsuarios, tablaUsuarios, "usuario", "idNeotel")

        cmbUsuarios.SelectedValue = WS_IDUSUARIO
        If cmbUsuarios.SelectedIndex > 0 Then
            cmbUsuarios.Enabled = False
        Else
            cmbUsuarios.SelectedIndex = 0
            cmbUsuarios.Enabled = True
        End If

        opc = 0
    End Sub

    Public Function ConectaNeotel(ByRef Usuario As String) As Boolean
        If Usuario = "" Then
            MsgBox("Debe ingresar usuario XLite", MsgBoxStyle.Exclamation, "Atención")
            txtusuxlite.Focus()
        Else

            vpPosicion = New clsSoapNeo
            If vpNeotel.CargarPosicion(Usuario) <> "" Then
                If vpPosicion.Usuario = Usuario Then
                    ConectaNeotel = True
                Else
                    MsgBox("El usuario XLite no corresponde", MsgBoxStyle.Exclamation, "Atención")
                End If
            End If
        End If
    End Function

    Private Sub frmLogin_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If opc = 0 Then
            End
        End If
    End Sub

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
        End
    End Sub

    Private Sub cmbUsuarios_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbUsuarios.SelectedIndexChanged
        If cmbUsuarios.SelectedIndex <= 0 Then
            txtusuxlite.Text = ""
            Exit Sub
        End If
        txtusuxlite.Text = cmbUsuarios.SelectedValue
        WS_IDUSUARIO = cmbUsuarios.Text
    End Sub
End Class