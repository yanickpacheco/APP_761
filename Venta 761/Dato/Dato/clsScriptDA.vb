Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports Entidad

Public Class clsScriptDA
    Dim con As New clsConexion 'SE CREA OBJETO DE CLSCONEXION PARA PODER CONECTARNOS A BASE DE DATOS
    Dim query As String 'VARIABLE QUE ALMACENARA EL PROCEDIMIENTO ALMACENDADO A CONSULTAR
    Dim tipoConexion As Integer 'VARIABLE QUE ALMACENARA LA BASE A LA CUAL CONSULTAR
    'VARIABLE QUE ALMACENARA LOS PARAMETROS A LA BASE DE DATOS
    Dim dt As New DataTable

    '**********METODO PARA BUSCAR SCRIPT PARA VISUALIZAR EN UN WEB BROWSER****************
    Public Function BuscarScriptPorIdTipoScript(ByVal _idCRM As Integer, ByVal _idTipoScript As Integer) As eScript


        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable
        Dim listTipoContrato As New List(Of eTipoContrato)


        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Gestion].[pa_ListaScriptPorId]"
        cmd.Parameters.AddWithValue("@idCRM", _idCRM)
        cmd.Parameters.AddWithValue("@idTipoScript", _idTipoScript)
        cmd.CommandText = sql
        dt = con.TraeDatos(cmd, 1)

        Dim row As New eScript

        For x As Int16 = 0 To dt.Rows.Count - 1

            row.idScript = IIf(IsDBNull(dt.Rows(x)("idScript")), Nothing, dt.Rows(x)("idScript"))
            row.contenidoScript = IIf(IsDBNull(dt.Rows(x)("contenidoScript")), Nothing, dt.Rows(x)("contenidoScript"))
            Row.idTipoScript = IIf(IsDBNull(dt.Rows(x)("idTipoScript")), Nothing, dt.Rows(x)("idTipoScript"))

        Next

        Return Row

    End Function


End Class
