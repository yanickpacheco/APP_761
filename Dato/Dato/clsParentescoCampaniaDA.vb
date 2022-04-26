Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports Entidad

Public Class clsParentescoCampaniaDA
    Dim con As New clsConexion 'SE CREA OBJETO DE CLSCONEXION PARA PODER CONECTARNOS A BASE DE DATOS
    Dim query As String 'VARIABLE QUE ALMACENARA EL PROCEDIMIENTO ALMACENDADO A CONSULTAR
    Dim tipoConexion As Integer 'VARIABLE QUE ALMACENARA LA BASE A LA CUAL CONSULTAR
    'VARIABLE QUE ALMACENARA LOS PARAMETROS A LA BASE DE DATOS
    Dim dt As New DataTable

    Public Function BuscaParentescoPorId(ByVal _idCRM As Integer, ByVal _tipoPersona As Integer) As List(Of eTipoParentesco)
        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable
        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Gestion].[pa_ListaParentescoPorCRM]"
        cmd.Parameters.AddWithValue("@CRM", _idCRM)
        cmd.Parameters.AddWithValue("@tipoPersona", _tipoPersona)
        cmd.CommandText = sql
        dt = con.TraeDatos(cmd, 1)

        Dim listParentesco As New List(Of eTipoParentesco)
        Dim parentesco As New eTipoParentesco
        parentesco.idParentesco = 0
        parentesco.nombreParentesco = "---No Ingresado---"
        listParentesco.Add(parentesco)

        For x As Int16 = 0 To dt.Rows.Count - 1
            Dim row As New eTipoParentesco
            row.idParentesco = IIf(IsDBNull(dt.Rows(x)("idParentesco")), Nothing, dt.Rows(x)("idParentesco"))
            row.nombreParentesco = IIf(IsDBNull(dt.Rows(x)("nombreParentesco")), Nothing, dt.Rows(x)("nombreParentesco"))
            'row.edadMin = IIf(IsDBNull(dt.Rows(x)("edadMin")), Nothing, dt.Rows(x)("edadMin"))
            'row.edadMax = IIf(IsDBNull(dt.Rows(x)("edadMax")), Nothing, dt.Rows(x)("edadMax"))
            listParentesco.Add(row)
        Next

        Return listParentesco

    End Function

    Public Function BuscaParentescoEdadPorId(ByVal _idCRM As Integer, ByVal _idTipoBeneficiarioCRM As Integer) As List(Of eTipoParentesco)
        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable

        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Gestion].[pa_ListaParentescoPorEdadCRM]"
        cmd.Parameters.AddWithValue("@CRM", _idCRM)
        'cmd.Parameters.AddWithValue("@tipoPersona", _tipoPersona)
        cmd.Parameters.AddWithValue("@idTipoBeneficiarioCRM", _idTipoBeneficiarioCRM)
        cmd.CommandText = sql
        dt = con.TraeDatos(cmd, 1)

        Dim listParentesco As New List(Of eTipoParentesco)
        Dim parentesco As New eTipoParentesco
        parentesco.idParentesco = 0
        parentesco.nombreParentesco = "---No Ingresado---"
        listParentesco.Add(parentesco)

        For x As Int16 = 0 To dt.Rows.Count - 1
            Dim row As New eTipoParentesco
            row.idPlan = IIf(IsDBNull(dt.Rows(x)("idPlan")), Nothing, dt.Rows(x)("idPlan"))
            row.idParentesco = IIf(IsDBNull(dt.Rows(x)("IdTipoBeneficiario")), Nothing, dt.Rows(x)("IdTipoBeneficiario"))
            row.nombreParentesco = IIf(IsDBNull(dt.Rows(x)("NombreTipoBeneficiario")), Nothing, dt.Rows(x)("NombreTipoBeneficiario"))
            row.edadMin = IIf(IsDBNull(dt.Rows(x)("edadDesde")), Nothing, dt.Rows(x)("edadDesde"))
            row.edadMax = IIf(IsDBNull(dt.Rows(x)("edadHasta")), Nothing, dt.Rows(x)("edadHasta"))
            listParentesco.Add(row)
        Next

        Return listParentesco

    End Function
End Class