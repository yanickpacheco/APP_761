Public Class frmEsperaDescanso
    Private Sub frmEsperaDescanso_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        asignaDescanso(idDescanso)

    End Sub

    Private Sub asignadescanso(ByVal id As Integer)
        'Baño
        'Break
        'Colacion
        'Capacitacion

        Select Case id
            Case 1
                lblNombreDescanso.Text = "BAÑO"
                pbImagenDescanso.Image = Global.APP_761.My.Resources.Resources.Baño
            Case 2
                lblNombreDescanso.Text = "BREAK"
                pbImagenDescanso.Image = Global.APP_761.My.Resources.Resources.Break
            Case 3
                lblNombreDescanso.Text = "COLACION"
                pbImagenDescanso.Image = Global.APP_761.My.Resources.Resources.Colacion
            Case 4
                lblNombreDescanso.Text = "CAPACITACION / RETROALIMENTACION"
                pbImagenDescanso.Image = Global.APP_761.My.Resources.Resources.Capacitacion

        End Select


    End Sub

    Private Sub btnFinDescanso_Click(sender As Object, e As EventArgs) Handles btnFinDescanso.Click
        vpPosicion.ActivarDescanso(vpPosicion.Usuario, False)
        IngresarDescanso(WS_IDUSUARIO, idDescanso, False)
        Me.Close()
    End Sub
End Class