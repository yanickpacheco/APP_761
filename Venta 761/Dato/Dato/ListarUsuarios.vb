Imports System.Data.SqlClient

Public Class clsListaUsuarioBaseDA
    Dim con As New clsConexion 'SE CREA OBJETO DE CLSCONEXION PARA PODER CONECTARNOS A BASE DE DATOS
    Dim query As String 'VARIABLE QUE ALMACENARA EL PROCEDIMIENTO ALMACENDADO A CONSULTAR
    Dim tipoConexion As Integer = 1 'VARIABLE QUE ALMACENARA LA BASE A LA CUAL CONSULTAR
    Dim dt As New DataTable

    Public Function ListarUsuarios(ByVal _crm As String) As DataTable
        Dim cmd As New SqlCommand
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = "Entidad.pa_ListaUsuarioCampaña"
        cmd.Parameters.AddWithValue("@crm", _crm)

        dt = con.TraeDatos(cmd, 1)

        Return dt
    End Function

End Class
