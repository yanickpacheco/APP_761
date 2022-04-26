Imports System.Data.SqlClient
Imports Entidad

Public Class clsMedioPagoDA

    Dim con As New clsConexion 'SE CREA OBJETO DE CLSCONEXION PARA PODER CONECTARNOS A BASE DE DATOS

    Public Function BuscaDatosMedioPago(ByVal _crm As Int32, ByVal _idCliente As Int32) As List(Of eMedioPago)
        Dim cmd As New SqlCommand
        Dim dt As New DataTable

        cmd.CommandType = CommandType.StoredProcedure

        cmd.CommandText = "[Config].[pa_ObtieneMedioDePago]"
        cmd.Parameters.AddWithValue("@crm", _crm)
        cmd.Parameters.AddWithValue("@idCliente", _idCliente)
        dt = con.TraeDatos(cmd, 1)

        Dim listMedioPago As New List(Of eMedioPago)
        For x As Int16 = 0 To dt.Rows.Count - 1
            Dim row As New eMedioPago

            row.idMedioPago = (IIf(IsDBNull(dt.Rows(x)("idMedioPago")), Nothing, dt.Rows(x)("idMedioPago")))
            row.Tarjeta = (IIf(IsDBNull(dt.Rows(x)("nombreMedioPago")), Nothing, dt.Rows(x)("nombreMedioPago")))
            row.NumeroMedioPago = (IIf(IsDBNull(dt.Rows(x)("numeroTarjeta")), Nothing, dt.Rows(x)("numeroTarjeta")))
            row.otroMedioPago = (IIf(IsDBNull(dt.Rows(x)("otroMedioPago")), Nothing, dt.Rows(x)("otroMedioPago")))

            listMedioPago.Add(row)
        Next

        Return listMedioPago
    End Function

    Public Function BuscaDatosMedioPagoPorId(ByVal _crm As Int32, ByVal _idCliente As Int32, ByVal _idMedioPago As Int32) As eMedioPago
        Dim cmd As New SqlCommand
        Dim dt As New DataTable

        cmd.CommandType = CommandType.StoredProcedure

        cmd.CommandText = "[Config].[pa_BuscaDatosMedioPagoPorId]"
        cmd.Parameters.AddWithValue("@crm", _crm)
        cmd.Parameters.AddWithValue("@idCliente", _idCliente)
        cmd.Parameters.AddWithValue("@idMedioPago", _idMedioPago)
        dt = con.TraeDatos(cmd, 1)

        Dim row As New eMedioPago
        For x As Int16 = 0 To dt.Rows.Count - 1
            row.idMedioPago = (IIf(IsDBNull(dt.Rows(x)("idMedioPago")), Nothing, dt.Rows(x)("idMedioPago")))
            row.Tarjeta = (IIf(IsDBNull(dt.Rows(x)("nombreMedioPago")), Nothing, dt.Rows(x)("nombreMedioPago")))
            row.NumeroMedioPago = (IIf(IsDBNull(dt.Rows(x)("numeroTarjeta")), Nothing, dt.Rows(x)("numeroTarjeta")))
            row.otroMedioPago = (IIf(IsDBNull(dt.Rows(x)("otroMedioPago")), Nothing, dt.Rows(x)("otroMedioPago")))
        Next

        Return row


    End Function

End Class
