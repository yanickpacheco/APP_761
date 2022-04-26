Imports System.Data.SqlClient
Imports Entidad
Public Class clsBeneficiarioDA
    Dim con As New clsConexion 'SE CREA OBJETO DE CLSCONEXION PARA PODER CONECTARNOS A BASE DE DATOS
    Dim query As String 'VARIABLE QUE ALMACENARA EL PROCEDIMIENTO ALMACENDADO A CONSULTAR
    Dim tipoConexion As Integer 'VARIABLE QUE ALMACENARA LA BASE A LA CUAL CONSULTAR
    ' 'VARIABLE QUE ALMACENARA LOS PARAMETROS A LA BASE DE DATOS
    Dim Tabla As New DataTable
    '**********************Netodo para eliminar beneficiario***********************************
    Public Sub Eliminar(ByVal id As String)
        Dim cmd As New SqlCommand
        cmd.CommandType = CommandType.StoredProcedure
        tipoConexion = 2
        cmd.CommandText = "[Gestion].[pa_Eliminar_Beneficiario]"
        cmd.Parameters.AddWithValue("@C_ID", id)
        con.Ejecutar(cmd, 2)
    End Sub
    '**********************Netodo para insertar beneficiario***********************************
    Public Sub Insertar(ByVal ben As eBeneficiario)
        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable
        Dim valido As Boolean


        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Gestion].[pa_IngresaBeneficiario]"
        cmd.Parameters.AddWithValue("@C_ID", ben.C_ID)
        cmd.Parameters.AddWithValue("@B_NRO", ben.B_Nro)
        cmd.Parameters.AddWithValue("@B_RUT", ben.B_Rut)
        cmd.Parameters.AddWithValue("@B_DV", ben.B_Dv)
        cmd.Parameters.AddWithValue("@B_NOMBRE", ben.B_Nombre)
        cmd.Parameters.AddWithValue("@B_PATERNO", ben.B_Paterno)
        cmd.Parameters.AddWithValue("@B_MATERNO", ben.B_Materno)
        cmd.Parameters.AddWithValue("@B_TIPO_BENEFICIARIO", ben.B_TipoBeneficiario)
        cmd.Parameters.AddWithValue("@B_NOMBRE_TIPO_BENEFICIARIO", ben.B_NOMBRE_PARENTESCO)
        cmd.Parameters.AddWithValue("@B_FECHA_NACIMIENTO", ben.B_FechaNacimiento)
        cmd.Parameters.AddWithValue("@B_PORCENTAJE", ben.B_Porcentaje)
        cmd.Parameters.AddWithValue("@B_SEXO", ben.B_Sexo)
        cmd.CommandText = sql
        valido = con.Ejecutar(cmd, 2)


    End Sub

    Public Function Carga_Beneficiarios(ByVal id As String) As List(Of eBeneficiario)
        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable

        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Gestion].[pa_ListaBeneficiarios]"
        cmd.CommandText = sql
        cmd.Parameters.AddWithValue("@C_ID", id)
        dt = con.TraeDatos(cmd, 2)

        Dim lstBeneficiarios As New List(Of eBeneficiario)

        For x As Int16 = 0 To dt.Rows.Count - 1
            Dim row As New eBeneficiario

            row.B_Rut = IIf(IsDBNull(dt.Rows(x)("B_Rut")), Nothing, dt.Rows(x)("B_Rut"))
            row.B_Dv = IIf(IsDBNull(dt.Rows(x)("B_Dv")), Nothing, dt.Rows(x)("B_Dv"))
            row.B_Nombre = IIf(IsDBNull(dt.Rows(x)("B_Nombre")), Nothing, dt.Rows(x)("B_Nombre"))
            row.B_Paterno = IIf(IsDBNull(dt.Rows(x)("B_Paterno")), Nothing, dt.Rows(x)("B_Paterno"))
            row.B_Materno = IIf(IsDBNull(dt.Rows(x)("B_Materno")), Nothing, dt.Rows(x)("B_Materno"))
            row.B_FechaNacimiento = IIf(IsDBNull(dt.Rows(x)("B_FechaNacimiento")), Nothing, dt.Rows(x)("B_FechaNacimiento"))
            row.B_Porcentaje = IIf(IsDBNull(dt.Rows(x)("B_Porcentaje")), Nothing, dt.Rows(x)("B_Porcentaje"))
            row.B_Sexo = IIf(IsDBNull(dt.Rows(x)("B_Sexo")), Nothing, dt.Rows(x)("B_Sexo"))
            row.B_TipoBeneficiario = IIf(IsDBNull(dt.Rows(x)("B_TipoBeneficiario")), Nothing, dt.Rows(x)("B_TipoBeneficiario"))


            lstBeneficiarios.Add(row)
        Next

        Return lstBeneficiarios
    End Function

End Class
