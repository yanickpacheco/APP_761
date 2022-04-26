Option Strict Off
Option Explicit On
Imports BI
Friend Class frmExclusiones
    Inherits System.Windows.Forms.Form



    Private Sub frmAnexos_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load


    End Sub

    Private Sub btnVolver_Click(sender As Object, e As EventArgs) Handles btnVolver.Click
        Me.Close()
    End Sub



    Friend WithEvents chkExclusiones As CheckedListBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents lblExclusiones As Label

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub btnIngresar_Click(sender As Object, e As EventArgs) Handles btnIngresar.Click


        If validadatos() = True Then

            Dim primero As Boolean = False
            For i As Integer = 0 To chkListBoxExclusiones.Items.Count - 1
                If (chkListBoxExclusiones.GetItemChecked(i)) Then
                    If primero = False Then
                        GESTION.G_Dato1 = chkListBoxExclusiones.Items(i).ToString
                        primero = True
                    Else
                        GESTION.G_Dato1 = GESTION.G_Dato1 & "|" & chkListBoxExclusiones.Items(i).ToString
                        primero = True
                    End If
                End If
            Next

            MsgBox("Gestiones ingresadas correctamente", vbInformation, "CORDIALPHONE")

        End If
    End Sub

    Private Function validadatos() As Boolean

        Dim valida As Boolean
        valida = False

        If chkListBoxExclusiones.CheckedItems.Count = 0 Then
            If MsgBox("No ha ingresado Exclusiones, desea continuar?", vbYesNoCancel + vbExclamation, "CordialPhone") = vbYes Then
                valida = True
            Else
                valida = False
            End If
        Else
            valida = True
        End If

        Return valida

    End Function

End Class