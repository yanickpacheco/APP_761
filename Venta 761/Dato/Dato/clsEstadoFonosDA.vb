Imports System.Collections.Generic
Imports Entidad

Public Class clsEstadoFonosDA
    'Dim con As New clsConexion 'SE CREA OBJETO DE CLSCONEXION PARA PODER CONECTARNOS A BASE DE DATOS
    'Dim query As String 'VARIABLE QUE ALMACENARA EL PROCEDIMIENTO ALMACENDADO A CONSULTAR
    'Dim tipoConexion As Integer 'VARIABLE QUE ALMACENARA LA BASE A LA CUAL CONSULTAR
    ' 'VARIABLE QUE ALMACENARA LOS PARAMETROS A LA BASE DE DATOS
    'Dim dt As New DataTable

    ''************Metodo para listar los estados de fono**************************************
    Public Function listarEstadoFono() As List(Of eEstadoFonos)
        '    Dim lstEdoFono As New List(Of eEstadoFonos)
        '    

        '    tipoConexion = 1
        '    vlSqlParam.Clear()
        '    query = "[Config].[pa_ListaEstadoFono]"

        '    dt = con.TraeDatosSinP(query, tipoConexion)
        '    For i As Int16 = 0 To dt.Rows.Count - 1

        '        Dim vlEdoFono As New eEstadoFonos

        '        vlEdoFono.efId = dt.Rows(i)("efId")
        '        vlEdoFono.efDescripcion = dt.Rows(i)("efDescripcion")
        '        vlEdoFono.efVigente = dt.Rows(i)("efVigente")

        '        lstEdoFono.Add(vlEdoFono)

        '    Next
        '    vlSqlParam.Clear()
        '    Return lstEdoFono
    End Function

    Public Sub guardaEstadoFono(ByVal _callID As String, ByVal _cal_id As String, ByVal _idEstadoCall As Int16)
        '    tipoConexion = 2
        '    vlSqlParam.Clear()

        '    query = "GestionVenta.pa_" + nomcampania + "_UpdateEstadoFonoCallID"
        '    vlSqlParam.Add("@callID", _callID, SqlDbType.VarChar)
        '    vlSqlParam.Add("@clave", _cal_id, SqlDbType.VarChar)
        '    vlSqlParam.Add("@idEstadoCall", _idEstadoCall, SqlDbType.Int)
        '    con.Ejecutar(query, vlSqlParam, tipoConexion)

        '    vlSqlParam.Clear()

    End Sub
End Class
