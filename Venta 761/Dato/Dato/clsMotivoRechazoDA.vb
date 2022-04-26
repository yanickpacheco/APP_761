Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports Entidad

Public Class clsMotivoRechazoDA

    Dim con As New clsConexion 'SE CREA OBJETO DE CLSCONEXION PARA PODER CONECTARNOS A BASE DE DATOS
    Dim query As String 'VARIABLE QUE ALMACENARA EL PROCEDIMIENTO ALMACENDADO A CONSULTAR
    Dim tipoConexion As Integer 'VARIABLE QUE ALMACENARA LA BASE A LA CUAL CONSULTAR
    'VARIABLE QUE ALMACENARA LOS PARAMETROS A LA BASE DE DATOS
    Dim Tabla As New DataTable

    Public Function BuscarMotivoRechazoPorSponsor(ByVal _idCrm As Integer) As List(Of eMotivoRechazo)
        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable
        Dim listTipoContrato As New List(Of eTipoContrato)


        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Entidad].[pa_ListaMotivosObjecionCRM]"
        cmd.Parameters.AddWithValue("@CRM", _idCrm)
        cmd.CommandText = sql
        dt = con.TraeDatos(cmd, 1)

        Dim lstMotivoRechazo As New List(Of eMotivoRechazo)

        If dt.Rows.Count > 0 Then

            Dim motivoRechazo As New eMotivoRechazo
            motivoRechazo.idMotivosObjecion = 0
            motivoRechazo.nombreMotivosObjecion = "---No Ingresado---"
            lstMotivoRechazo.Add(motivoRechazo)

            For x As Int16 = 0 To dt.Rows.Count - 1
                Dim row As New eMotivoRechazo
                row.idMotivosObjecion = IIf(IsDBNull(dt.Rows(x)("IdMotivosObjecion")), Nothing, dt.Rows(x)("IdMotivosObjecion"))
                row.nombreMotivosObjecion = IIf(IsDBNull(dt.Rows(x)("nombreMotivosObjecion")), Nothing, dt.Rows(x)("nombreMotivosObjecion"))
                lstMotivoRechazo.Add(row)
            Next

        Else
            MsgBox("Campaña no tiene los Motivos de Rechazo Configurados. Revisar con Supervisor", vbExclamation, "CORDIALPHONE")
        End If

        Return lstMotivoRechazo

    End Function


    Public Function BuscaMotivoRechazoCampaniaPorId(ByVal _idMotivosObjecion As Int32) As eMotivoRechazoCampania
        'vlSqlParam.Clear()
        'Tabla.Clear()

        'tipoConexion = 1
        'query = "Config.[pa_BuscaMotivoRechazoCampaniaPorId]"
        'vlSqlParam.Add("@idMotivosObjecionCampania", _idMotivosObjecion, SqlDbType.Int)

        'Tabla = con.TraeDatosConP(vlSqlParam, query, tipoConexion)
        ''Dim ListData As New List(Of eMotivoRechazoCampania)
        'Dim row As New eMotivoRechazoCampania
        'For x As Int16 = 0 To Tabla.Rows.Count - 1

        '    row.idMotivosObjecionCampania = IIf(IsDBNull(Tabla.Rows(x)("idMotivosObjecionCampania")), Nothing, Tabla.Rows(x)("idMotivosObjecionCampania"))
        '    row.idMotivosObjecion = IIf(IsDBNull(Tabla.Rows(x)("idMotivosObjecion")), Nothing, Tabla.Rows(x)("idMotivosObjecion"))
        '    row.espCodigo = IIf(IsDBNull(Tabla.Rows(x)("espCodigo")), Nothing, Tabla.Rows(x)("espCodigo"))
        '    row.calcodigo = IIf(IsDBNull(Tabla.Rows(x)("calcodigo")), Nothing, Tabla.Rows(x)("calcodigo"))
        '    row.otro = IIf(IsDBNull(Tabla.Rows(x)("otro")), Nothing, Tabla.Rows(x)("otro"))
        '    'ListData.Add(row)
        'Next

        'Return row

    End Function

End Class
