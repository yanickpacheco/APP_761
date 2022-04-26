Imports Dato
Imports Entidad

Public Class clsRestriccionBI

    Public Function BuscarRestriccionPorIdPlan(ByVal _idPlan As Int64) As List(Of eRestriccion)
        Dim daRestriccion As New clsRestriccionDA
        Return daRestriccion.BuscarRestriccionPorIdPlan(_idPlan)
    End Function

End Class
