Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports Entidad

Public Class clsTipoContratoDA
    Dim con As New clsConexion 'SE CREA OBJETO DE CLSCONEXION PARA PODER CONECTARNOS A BASE DE DATOS
    Dim query As String 'VARIABLE QUE ALMACENARA EL PROCEDIMIENTO ALMACENDADO A CONSULTAR
    Dim tipoConexion As Integer 'VARIABLE QUE ALMACENARA LA BASE A LA CUAL CONSULTAR
    'VARIABLE QUE ALMACENARA LOS PARAMETROS A LA BASE DE DATOS
    Dim Tabla As New DataTable

    '**************METODO PARA LISTAR LOS PLANES POR CAMPAÑA**********************
    Public Function ListaTipoContratoPorCampania(ByVal _idCRM As Int32) As List(Of eTipoContrato)
        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable
        Dim listTipoContrato As New List(Of eTipoContrato)


        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Entidad].[pa_ListaTipoContratoPorCRM]"
        cmd.Parameters.AddWithValue("@CRM", _idCRM)
        cmd.CommandText = sql
        dt = con.TraeDatos(cmd, 1)

        If dt.Rows.Count > 0 Then

            Dim tipoContrato As New eTipoContrato

            tipoContrato.idTipoContrato = 0
            tipoContrato.nombreTipoContrato = "---No ingresado---"
            listTipoContrato.Add(tipoContrato)

            For x As Int16 = 0 To dt.Rows.Count - 1
                Dim row As New eTipoContrato
                row.idTipoContrato = IIf(IsDBNull(dt.Rows(x)("idTipoContrato")), Nothing, dt.Rows(x)("idTipoContrato"))
                row.nombreTipoContrato = IIf(IsDBNull(dt.Rows(x)("nombreTipoContrato")), Nothing, dt.Rows(x)("nombreTipoContrato"))
                row.cantidadBeneficiarios = IIf(IsDBNull(dt.Rows(x)("cantidadBeneficiarios")), Nothing, dt.Rows(x)("cantidadBeneficiarios"))
                row.cantidadAdicionales = IIf(IsDBNull(dt.Rows(x)("cantidadAdicionales")), Nothing, dt.Rows(x)("cantidadAdicionales"))
                listTipoContrato.Add(row)
            Next

        End If

        Return listTipoContrato

    End Function

    '**************METODO PARA BUSCAR LOS TIPOS DE CONTRATO POR ID DE CONTARTO***********************
    Public Function BuscarTipoContratoPorIdTipoContrato(ByVal _idTipoContrato As Int16) As eTipoContrato
        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable
        Dim listTipoContrato As New List(Of eTipoContrato)


        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Entidad].[pa_ListaTipoContratoPorId]"
        cmd.Parameters.AddWithValue("@idTipoContrato", _idTipoContrato)
        cmd.CommandText = sql
        dt = con.TraeDatos(cmd, 1)

        Dim row As New eTipoContrato

        If dt.Rows.Count > 0 Then
            For x As Int16 = 0 To dt.Rows.Count - 1
                row.idTipoContrato = IIf(IsDBNull(dt.Rows(x)("idTipoContrato")), Nothing, dt.Rows(x)("idTipoContrato"))
                row.nombreTipoContrato = IIf(IsDBNull(dt.Rows(x)("nombreTipoContrato")), Nothing, dt.Rows(x)("nombreTipoContrato"))
                row.cantidadBeneficiarios = IIf(IsDBNull(dt.Rows(x)("cantidadBeneficiarios")), Nothing, dt.Rows(x)("cantidadBeneficiarios"))
                row.cantidadAdicionales = IIf(IsDBNull(dt.Rows(x)("cantidadAdicionales")), Nothing, dt.Rows(x)("cantidadAdicionales"))
            Next
        End If

        Return Row

    End Function

    '**************METODO PARA LISTAR LOS PLANES POR CAMPAÑA**********************
    Public Function ListaTipoContratoPosAdicionales(ByVal _calCodigo As Int32, ByVal _cantAdic As SByte) As List(Of eTipoContrato)
        'Dim listTipoContrato As New List(Of eTipoContrato)
        'vlSqlParam.Clear()
        'Tabla.Clear()

        'tipoConexion = 1
        'query = "dbo.pa_ListaTipoContratoPosAdicionales"
        'vlSqlParam.Add("@calcodigo", _calCodigo, SqlDbType.Int)
        'vlSqlParam.Add("@CantAdicionales", _cantAdic, SqlDbType.Int)

        'Tabla = con.TraeDatosConP(vlSqlParam, query, tipoConexion)

        'For x As Int16 = 0 To Tabla.Rows.Count - 1
        '    Dim row As New eTipoContrato
        '    row.idTipoContrato = IIf(IsDBNull(Tabla.Rows(x)("idTipoContrato")), Nothing, Tabla.Rows(x)("idTipoContrato"))
        '    row.nombreTipoContrato = IIf(IsDBNull(Tabla.Rows(x)("nombreTipoContrato")), Nothing, Tabla.Rows(x)("nombreTipoContrato"))
        '    listTipoContrato.Add(row)
        'Next

        'Return listTipoContrato

    End Function

End Class
