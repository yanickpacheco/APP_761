Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports Dato
Imports Entidad

Public Class clsCampoAdicionalDA

    Dim con As New clsConexion 'SE CREA OBJETO DE CLSCONEXION PARA PODER CONECTARNOS A BASE DE DATOS
    Dim query As String 'VARIABLE QUE ALMACENARA EL PROCEDIMIENTO ALMACENDADO A CONSULTAR
    Dim tipoConexion As Integer 'VARIABLE QUE ALMACENARA LA BASE A LA CUAL CONSULTAR
    '  'VARIABLE QUE ALMACENARA LOS PARAMETROS A LA BASE DE DATOS
    Dim dt As New DataTable

    '********* Busca todos los campos adicionales de la campaña filtrada ***************
    Public Function BuscarCampoAdicionalPorCalCodigo(ByVal _idCrm As Int32) As List(Of eCampoAdicional)

        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable


        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Config].[pa_BuscarCampoAdicionalPorCRM]"
        cmd.Parameters.AddWithValue("@idCRM", _idCrm)

        cmd.CommandText = sql
        dt = con.TraeDatos(cmd, 1)

        Dim listCampoAdicional As New List(Of eCampoAdicional)

        For x As Int16 = 0 To dt.Rows.Count - 1
            Dim row As New eCampoAdicional

            row.idCampoAdicional = IIf(IsDBNull(dt.Rows(x)("idCampoAdicional")), Nothing, dt.Rows(x)("idCampoAdicional"))
            row.nombreCampoAdicional = IIf(IsDBNull(dt.Rows(x)("nombreCampoAdicional")), Nothing, dt.Rows(x)("nombreCampoAdicional"))
            row.valorCampo = IIf(IsDBNull(dt.Rows(x)("valorCampo")), Nothing, dt.Rows(x)("valorCampo"))
            row.idCRM = IIf(IsDBNull(dt.Rows(x)("idcrm")), Nothing, dt.Rows(x)("idcrm"))
            row.fechaCreacion = IIf(IsDBNull(dt.Rows(x)("fechaCreacion")), Nothing, dt.Rows(x)("fechaCreacion"))
            row.orden = IIf(IsDBNull(dt.Rows(x)("orden")), Nothing, dt.Rows(x)("orden"))
            row.activo = IIf(IsDBNull(dt.Rows(x)("activo")), Nothing, dt.Rows(x)("activo"))
            listCampoAdicional.Add(row)
        Next


        Return listCampoAdicional

    End Function

    '********* BUSCA TODOS LOS DATOS CONFIGURADOS EN LA TABLA CAMPOADICIONAL ***************
    Public Function BuscaDatosCampoAdicional(ByVal _idCrm As Int32, ByVal _idCliente As Int32) As List(Of eCampoAdicional)

        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable


        cmd.CommandType = CommandType.StoredProcedure

        sql = "Config.pa_BuscaDatosCampoAdicional"
        cmd.Parameters.AddWithValue("@idCRM", _idCrm)
        cmd.Parameters.AddWithValue("@idCliente", _idCliente)

        cmd.CommandText = sql
        dt = con.TraeDatos(cmd, 1)

        Dim listCampoAdicional As New List(Of eCampoAdicional)

        For x As Int16 = 0 To dt.Rows.Count - 1
            Dim row As New eCampoAdicional

            row.nombreCampoAdicional = (IIf(IsDBNull(dt.Rows(x)("nombreCampo")), Nothing, dt.Rows(x)("nombreCampo")))
            row.valorCampo = (IIf(IsDBNull(dt.Rows(x)("camposTabla")), Nothing, dt.Rows(x)("camposTabla")))
            listCampoAdicional.Add(row)
        Next


        Return listCampoAdicional

    End Function


End Class
