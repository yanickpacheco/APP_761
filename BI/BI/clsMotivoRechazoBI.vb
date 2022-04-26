Imports Dato
Imports Entidad

Public Class clsMotivoRechazoBI


    Public Function BuscarMotivoRechazoPorSponsor(ByVal idCRM As Integer) As List(Of eMotivoRechazo)
        Dim daMotivoRechazo As New clsMotivoRechazoDA
        Return daMotivoRechazo.BuscarMotivoRechazoPorSponsor(idCRM)
    End Function

    Public Function BuscaMotivoRechazoCampaniaPorId(ByVal _idMotivosObjecion As Int32) As eMotivoRechazoCampania
        Dim daMotivoRechazo As New clsMotivoRechazoDA
        Return daMotivoRechazo.BuscaMotivoRechazoCampaniaPorId(_idMotivosObjecion)
    End Function

End Class
