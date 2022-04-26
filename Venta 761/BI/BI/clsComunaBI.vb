Imports Dato
Imports Entidad
Public Class clsComunaBI

    Public Function listarComuna() As List(Of eComuna)
        Dim daComuna As New clsComunaDA
        Return daComuna.listarComuna()
    End Function

End Class
