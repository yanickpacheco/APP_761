Public Class frmDescanso
    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub

    Private Sub btnActivarDescanso_Click(sender As Object, e As EventArgs) Handles btnActivarDescanso.Click
        If validadatos() = True Then

            Dim idNeotel As String = ""

            If rdbDescanso1.Checked = True Then
                idDescanso = 1
                idNeotel = "*10"
            ElseIf rdbDescanso2.Checked = True Then
                idDescanso = 2
                idNeotel = "*03"
            ElseIf rdbDescanso3.Checked = True Then
                idDescanso = 3
                idNeotel = "*23"
            ElseIf rdbDescanso4.Checked = True Then
                idDescanso = 4
                idNeotel = "*07"
            End If
            If idNeotel <> "" Then
                Me.Cursor = Cursors.WaitCursor
                vpPosicion.Discar(vpPosicion.Usuario, idNeotel)
                IngresarDescanso(WS_IDUSUARIO, idDescanso, True)
                Me.Cursor = Cursors.Default
                MsgBox("Descanso Activado", vbInformation, "CORDIAL PHONE")
                frmVenta.tmrPostview.Enabled = False
                frmEsperaDescanso.ShowDialog()
                frmVenta.tmrPostview.Enabled = True
                Me.Close()
            End If

        End If

    End Sub

    Private Function validadatos() As Boolean
        validadatos = False

        If rdbDescanso1.Checked = False And rdbDescanso2.Checked = False And rdbDescanso3.Checked = False And rdbDescanso4.Checked = False Then
            MsgBox("Debe ingresar al menos una opcion para Descanso. Si no desea ninguna, oprima Salir ", vbInformation, "Descanso Valido")
            rdbDescanso1.Focus()
            Exit Function
        End If

        validadatos = True
    End Function

End Class