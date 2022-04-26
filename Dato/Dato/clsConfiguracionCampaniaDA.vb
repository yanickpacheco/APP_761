Imports System.Data.SqlClient
Imports Entidad

Public Class clsConfiguracionCampaniaDA




    Dim cnn As New SqlConnection("data source = MAURICIO-GOMEZ; initial catalog = SOPORTE; User Id= mg; Password=Cordial.passs")

    Dim con As New clsConexion 'SE CREA OBJETO DE CLSCONEXION PARA PODER CONECTARNOS A BASE DE DATOS
    Dim query As String 'VARIABLE QUE ALMACENARA EL PROCEDIMIENTO ALMACENDADO A CONSULTAR
    Dim tipoConexion As Integer 'VARIABLE QUE ALMACENARA LA BASE A LA CUAL CONSULTAR
    '  'VARIABLE QUE ALMACENARA LOS PARAMETROS A LA BASE DE DATOS
    Dim dt As New DataTable



    Public Function ObtenerConfiguracionCampania(ByRef _CRM As String) As eProyecto

        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable


        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Entidad].[pa_ListarCampañas]"
        cmd.Parameters.AddWithValue("@CRM", _CRM)
        cmd.CommandText = sql
        dt = con.TraeDatos(cmd, 1)

        Dim row As New eProyecto
        For x As Int16 = 0 To dt.Rows.Count - 1

            row.idCRM = IIf(dt.Rows(x)("CRM") Is DBNull.Value, Nothing, dt.Rows(x)("CRM"))
            row.IdTelefonia = IIf(dt.Rows(x)("IdTelefonia") Is DBNull.Value, Nothing, dt.Rows(x)("IdTelefonia"))
            row.nombre = IIf(dt.Rows(x)("nombreCRM") Is DBNull.Value, Nothing, dt.Rows(x)("nombreCRM"))
            row.ServidorBDD = IIf(dt.Rows(x)("servidorBDD") Is DBNull.Value, Nothing, dt.Rows(x)("servidorBDD"))

        Next

        Return row

    End Function

End Class

