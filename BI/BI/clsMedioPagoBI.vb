Imports Entidad
Imports Dato

Public Class clsMedioPagoBI
    Dim daMedioPago As New clsMedioPagoDA

    Public Function BuscaDatosMedioPago(ByVal _crm As Int32, ByVal _idCliente As Int32) As List(Of eMedioPago)
        Return daMedioPago.BuscaDatosMedioPago(_crm, _idCliente)
    End Function

    Public Function BuscaDatosMedioPagoPorId(ByVal _crm As Int32, ByVal _idCliente As Int32, ByVal _idMedioPago As Int32) As eMedioPago
        Return daMedioPago.BuscaDatosMedioPagoPorId(_crm, _idCliente, _idMedioPago)
    End Function

End Class
