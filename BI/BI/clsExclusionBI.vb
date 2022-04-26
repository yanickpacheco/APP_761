Imports Dato
Imports Entidad

Public Class clsExclusionBI

    Public Function listarExclusiones(ByVal idCRM As Integer) As List(Of eExclusion)
        Dim daExclusion As New clsExclusionDA
        Return daExclusion.listarExclusiones(idCRM)
    End Function

End Class
