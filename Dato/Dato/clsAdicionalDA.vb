Imports System.Data.SqlClient
Imports Entidad
Public Class clsAdicionalDA
    Dim con As New clsConexion 'SE CREA OBJETO DE CLSCONEXION PARA PODER CONECTARNOS A BASE DE DATOS

    Public Sub Insertar(ByVal adic As eAdicional)
        Dim cmd As New SqlCommand
        Dim dt As New DataTable
        Dim valido As Boolean

        cmd.CommandType = CommandType.StoredProcedure

        cmd.CommandText = "[Gestion].[pa_IngresaAdicional]"
        cmd.Parameters.AddWithValue("@C_ID", adic.C_ID)
        cmd.Parameters.AddWithValue("@A_NRO", adic.A_NRO)
        cmd.Parameters.AddWithValue("@A_RUT", adic.A_Rut)
        cmd.Parameters.AddWithValue("@A_DV", adic.A_Dv)
        cmd.Parameters.AddWithValue("@A_NOMBRE", adic.A_Nombre)
        cmd.Parameters.AddWithValue("@A_PATERNO", adic.A_Paterno)
        cmd.Parameters.AddWithValue("@A_MATERNO", adic.A_Materno)
        cmd.Parameters.AddWithValue("@A_FECHA_NACIMIENTO", adic.A_FechaNacimiento)
        cmd.Parameters.AddWithValue("@A_SEXO", adic.A_Sexo)
        cmd.Parameters.AddWithValue("@A_TIPO_BENEFICIARIO", adic.A_TipoBeneficiario)
        cmd.Parameters.AddWithValue("@A_NombreBeneficiario", adic.A_NombreBeneficiario)
        cmd.Parameters.AddWithValue("@a_primaUf", adic.a_primaUf)
        cmd.Parameters.AddWithValue("@a_primaPesos", adic.a_primaPesos)
        cmd.Parameters.AddWithValue("@a_salud", adic.a_salud)
        cmd.Parameters.AddWithValue("@idPlanAdic", adic.idPlanAdic)

        valido = con.Ejecutar(cmd, 2)
    End Sub

    Public Sub Eliminar(ByVal _c_id As Int64)
        Dim cmd As New SqlCommand
        Dim dt As New DataTable

        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = "gestion.pa_EliminaAdicional"
        cmd.Parameters.AddWithValue("@c_id", _c_id)

        con.Ejecutar(cmd, 2)
    End Sub
    Public Function Carga_Adicional(ByVal id As String) As List(Of eAdicional)
        Dim cmd As New SqlCommand
        Dim dt As New DataTable

        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = "[Gestion].[pa_ListaAdicionales]"
        cmd.Parameters.AddWithValue("@C_ID", id)
        dt = con.TraeDatos(cmd, 2)

        Dim lstAdicional As New List(Of eAdicional)

        For x As Int16 = 0 To dt.Rows.Count - 1
            Dim row As New eAdicional
            row.A_Rut = IIf(IsDBNull(dt.Rows(x)("A_Rut")), Nothing, dt.Rows(x)("A_Rut"))
            row.A_Dv = IIf(IsDBNull(dt.Rows(x)("A_Dv")), Nothing, dt.Rows(x)("A_Dv"))
            row.A_Nombre = IIf(IsDBNull(dt.Rows(x)("A_Nombre")), Nothing, dt.Rows(x)("A_Nombre"))
            row.A_Paterno = IIf(IsDBNull(dt.Rows(x)("A_Paterno")), Nothing, dt.Rows(x)("A_Paterno"))
            row.A_Materno = IIf(IsDBNull(dt.Rows(x)("A_Materno")), Nothing, dt.Rows(x)("A_Materno"))
            row.A_FechaNacimiento = IIf(IsDBNull(dt.Rows(x)("A_FechaNacimiento")), Nothing, dt.Rows(x)("A_FechaNacimiento"))
            row.A_Sexo = IIf(IsDBNull(dt.Rows(x)("A_Sexo")), Nothing, dt.Rows(x)("A_Sexo"))
            row.A_TipoBeneficiario = IIf(IsDBNull(dt.Rows(x)("A_TipoBeneficiario")), Nothing, dt.Rows(x)("A_TipoBeneficiario"))
            row.A_NombreBeneficiario = IIf(IsDBNull(dt.Rows(x)("A_NombreBeneficiario")), Nothing, dt.Rows(x)("A_NombreBeneficiario"))
            row.a_primaUf = IIf(IsDBNull(dt.Rows(x)("a_primaUf")), Nothing, dt.Rows(x)("a_primaUf"))
            row.idPlanAdic = IIf(IsDBNull(dt.Rows(x)("idPlanAdic")), Nothing, dt.Rows(x)("idPlanAdic"))
            row.a_salud = IIf(IsDBNull(dt.Rows(x)("a_salud")), Nothing, dt.Rows(x)("a_salud"))
            row.a_primaPesos = IIf(IsDBNull(dt.Rows(x)("a_primaPesos")), Nothing, dt.Rows(x)("a_primaPesos"))

            lstAdicional.Add(row)
        Next

        Return lstAdicional

    End Function

End Class
