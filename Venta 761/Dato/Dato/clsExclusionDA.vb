Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports Entidad
Public Class clsExclusionDA
    Dim con As New clsConexion 'SE CREA OBJETO DE CLSCONEXION PARA PODER CONECTARNOS A BASE DE DATOS
    Dim query As String 'VARIABLE QUE ALMACENARA EL PROCEDIMIENTO ALMACENDADO A CONSULTAR
    Dim tipoConexion As Integer 'VARIABLE QUE ALMACENARA LA BASE A LA CUAL CONSULTAR
    'VARIABLE QUE ALMACENARA LOS PARAMETROS A LA BASE DE DATOS
    Dim dt As New DataTable

    '****************Metodo para listar todas las comunas**********************************
    ''' <summary>
    ''' Metodo para listar todas las Exclusiones
    ''' </summary>
    ''' <returns>List(Of clsComunaEnt)</returns>
    ''' <remarks></remarks>
    Public Function listarExclusiones(ByVal idCRM As Integer) As List(Of eExclusion)

        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable


        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Entidad].[pa_listaExclusiones]"
        cmd.Parameters.AddWithValue("@idCRM", idCRM)
        cmd.CommandText = sql
        dt = con.TraeDatos(cmd, 1)

        Dim listExclusiones As New List(Of eExclusion)

        For x As Int16 = 0 To dt.Rows.Count - 1
            Dim row As New eExclusion

            row.IdExclusion = IIf(IsDBNull(dt.Rows(x)("idExclusion")), Nothing, dt.Rows(x)("idExclusion"))
            row.Exclusion = IIf(IsDBNull(dt.Rows(x)("Exclusion")), Nothing, dt.Rows(x)("Exclusion"))
            listExclusiones.Add(row)
        Next

        Return listExclusiones

    End Function
End Class
