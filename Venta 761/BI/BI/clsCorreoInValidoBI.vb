Imports Dato
Imports Entidad
Public Class clsCorreoInvalidoBI

    Public Function listarCorreosInvalido() As List(Of eCorreoInvalido)
        Dim daCorreoInv As New clsCorreoInvalidoDA
        Return daCorreoInv.listarCorreosInvalidos
    End Function

End Class
