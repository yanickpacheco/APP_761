Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports Entidad
Public Class clsUsuarioDA
    Dim con As New clsConexion 'SE CREA OBJETO DE CLSCONEXION PARA PODER CONECTARNOS A BASE DE DATOS
    Dim query As String 'VARIABLE QUE ALMACENARA EL PROCEDIMIENTO ALMACENDADO A CONSULTAR
    Dim tipoConexion As Integer 'VARIABLE QUE ALMACENARA LA BASE A LA CUAL CONSULTAR
    Dim tabla As New DataTable


    Function FechaServidor() As DateTime
        'Dim fecha As DateTime
        'fecha = Nothing
        'tipoConexion = 1

        'query = "[dbo].[Fecha_Servidor]"
        'Tabla = con.TraeDatosSinP(query, tipoConexion)

        'For Each row As DataRow In Tabla.Rows
        '    fecha = row.Item(0)
        'Next

        'Return fecha
    End Function

    ''' <summary>
    ''' Obtiene una usuario filtrado por usuario
    ''' </summary>
    ''' <param name="user"></param>
    ''' <returns>Entidad usuario</returns>
    ''' <remarks></remarks>
    Public Function Validar_user(ByVal user As String) As Boolean

        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable

        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Entidad].[pa_validaUsuario]"
        cmd.Parameters.AddWithValue("@usuario", user)
        cmd.CommandText = sql
        dt = con.TraeDatos(cmd, 2)

        Dim venta As Boolean
        venta = False

        If dt.Rows.Count <= 0 Then
            Return venta
        Else
            venta = True
        End If
        Return venta

    End Function

    ''' <summary>
    ''' Indica si usuario es regrabador
    ''' </summary>
    ''' <param name="user"></param>
    ''' <returns>Entidad usuario</returns>
    ''' <remarks></remarks>
    Public Function Validar_REG(ByVal user As String) As Boolean

        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable

        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Entidad].[pa_validaUsuario_Reg]"
        cmd.Parameters.AddWithValue("@usuario", user)
        cmd.CommandText = sql
        dt = con.TraeDatos(cmd, 2)

        Dim venta As Boolean
        venta = False

        If dt.Rows.Count <= 0 Then
            Return venta
        Else
            venta = True
        End If
        Return venta

    End Function
End Class
