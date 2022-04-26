Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports Entidad
Public Class clsCiudadDA
    Dim con As New clsConexion 'SE CREA OBJETO DE CLSCONEXION PARA PODER CONECTARNOS A BASE DE DATOS
    Dim query As String 'VARIABLE QUE ALMACENARA EL PROCEDIMIENTO ALMACENDADO A CONSULTAR
    Dim tipoConexion As Integer 'VARIABLE QUE ALMACENARA LA BASE A LA CUAL CONSULTAR
    'VARIABLE QUE ALMACENARA LOS PARAMETROS A LA BASE DE DATOS
    Dim dt As New DataTable


    ''' <summary>
    ''' Obtiene una ciudad filtrado por idCiudad
    ''' </summary>
    ''' <param name="_idCiudad"></param>
    ''' <returns>Entidad ciudad</returns>
    ''' <remarks></remarks>
    Public Function BuscaCiudadPorIdCiudad(ByVal _idCiudad As Int32) As eCiudad
        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable


        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Entidad].[pa_buscarCiudadPorId]"
        cmd.CommandText = sql
        cmd.Parameters.AddWithValue("@idCiudad", _idCiudad)
        dt = con.TraeDatos(cmd, 1)

        Dim row As New eCiudad

        For x As Int16 = 0 To dt.Rows.Count - 1
            row.idCiudad = IIf(IsDBNull(dt.Rows(x)("idCiudad")), Nothing, dt.Rows(x)("idCiudad"))
            row.idRegion = IIf(IsDBNull(dt.Rows(x)("idRegion")), Nothing, dt.Rows(x)("idRegion"))
            row.nombreCiudad = IIf(IsDBNull(dt.Rows(x)("nombreCiudad")), Nothing, dt.Rows(x)("nombreCiudad"))
        Next

        Return row

    End Function

    Public Function ListaCiudad() As List(Of eCiudad)
        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable


        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Entidad].[pa_listarCiudad]"
        cmd.CommandText = sql
        dt = con.TraeDatos(cmd, 1)

        Dim list As New List(Of eCiudad)
        Dim row As New eCiudad
        row.idCiudad = 0
        row.nombreCiudad = "--- Sin selección ---"
        list.Add(row)

        For x As Int16 = 0 To dt.Rows.Count - 1
            row = New eCiudad
            row.idCiudad = IIf(IsDBNull(dt.Rows(x)("idCiudad")), Nothing, dt.Rows(x)("idCiudad"))
            row.idRegion = IIf(IsDBNull(dt.Rows(x)("idRegion")), Nothing, dt.Rows(x)("idRegion"))
            row.nombreCiudad = IIf(IsDBNull(dt.Rows(x)("nombreCiudad")), Nothing, dt.Rows(x)("nombreCiudad"))
            list.Add(row)
        Next

        Return list

    End Function




End Class
