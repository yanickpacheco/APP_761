Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports Entidad

Public Class clsPlanDA
    Dim con As New clsConexion 'SE CREA OBJETO DE CLSCONEXION PARA PODER CONECTARNOS A BASE DE DATOS
    Dim query As String 'VARIABLE QUE ALMACENARA EL PROCEDIMIENTO ALMACENDADO A CONSULTAR
    Dim tipoConexion As Integer 'VARIABLE QUE ALMACENARA LA BASE A LA CUAL CONSULTAR
    'VARIABLE QUE ALMACENARA LOS PARAMETROS A LA BASE DE DATOS
    Dim Tabla As New DataTable

    '******************************METODO PARA BUSCRA LOS PLANES POR TIPO DE CONTRATO************
    Public Function BuscarPlanPorTipoContrato(ByVal _idTipoContrato As Int16, ByVal _idCRM As Int32) As List(Of ePlan)

        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable
        Dim listPlan As New List(Of ePlan)


        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Entidad].[pa_BuscaPlanPorTipoContrato]"
        cmd.Parameters.AddWithValue("@idTipoContrato", _idTipoContrato)
        cmd.Parameters.AddWithValue("@CRM", _idCRM)
        cmd.CommandText = sql
        dt = con.TraeDatos(cmd, 1)

        If dt.Rows.Count > 0 Then

            Dim plan As New ePlan
            plan.idPlan = 0
            plan.descripcionPlan = "---No ingresado---"
            listPlan.Add(plan)

            For x As Int16 = 0 To dt.Rows.Count - 1
                Dim row As New ePlan
                row.idPlan = IIf(IsDBNull(dt.Rows(x)("idPlan")), Nothing, dt.Rows(x)("idPlan"))
                row.primaUF = IIf(IsDBNull(dt.Rows(x)("primaUF")), Nothing, dt.Rows(x)("primaUF"))
                row.PrimaPesos = IIf(IsDBNull(dt.Rows(x)("PrimaPesos")), Nothing, dt.Rows(x)("PrimaPesos"))
                row.idTipoContrato = IIf(IsDBNull(dt.Rows(x)("idTipoContrato")), Nothing, dt.Rows(x)("idTipoContrato"))
                row.descripcionPlan = IIf(IsDBNull(dt.Rows(x)("descripcionPlan")), Nothing, dt.Rows(x)("descripcionPlan"))
                row.idCRM = IIf(IsDBNull(dt.Rows(x)("idCRM")), Nothing, dt.Rows(x)("idCRM"))
                row.activo = IIf(IsDBNull(dt.Rows(x)("activo")), Nothing, dt.Rows(x)("activo"))
                listPlan.Add(row)
            Next

        End If
        Return listPlan


    End Function


    '**********************METODO PARA BUSCAR LOS PLANES POR ID DE PLAN*****************************************
    Public Function BuscarPlanPorIdPlan(ByVal _idPlan As Int64) As ePlan

        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable
        Dim listPlan As New List(Of ePlan)


        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Entidad].[pa_BuscaPlanPorIdPlan]"
        cmd.Parameters.AddWithValue("@idPlan", _idPlan)
        cmd.CommandText = sql
        dt = con.TraeDatos(cmd, 1)

        Dim row As New ePlan

        For x As Int16 = 0 To dt.Rows.Count - 1

            row.idPlan = IIf(IsDBNull(dt.Rows(x)("idPlan")), Nothing, dt.Rows(x)("idPlan"))
            row.primaUF = IIf(IsDBNull(dt.Rows(x)("primaUF")), Nothing, dt.Rows(x)("primaUF"))
            row.PrimaPesos = IIf(IsDBNull(dt.Rows(x)("PrimaPesos")), Nothing, dt.Rows(x)("PrimaPesos"))
            row.idTipoContrato = IIf(IsDBNull(dt.Rows(x)("idTipoContrato")), Nothing, dt.Rows(x)("idTipoContrato"))
            row.descripcionPlan = IIf(IsDBNull(dt.Rows(x)("descripcionPlan")), Nothing, dt.Rows(x)("descripcionPlan"))
            row.idCRM = IIf(IsDBNull(dt.Rows(x)("idCRM")), Nothing, dt.Rows(x)("idCRM"))
            row.activo = IIf(IsDBNull(dt.Rows(x)("activo")), Nothing, dt.Rows(x)("activo"))

        Next
        Return row

    End Function
    '******************METODO PARA LISTAR LOS PLANES ADICIONALES*******************************
    Public Function ListaPlanAdicionales(ByVal _idTipoContrato As eTipoContrato, ByVal _calCodigo As Int32) As List(Of ePlan)
        'Dim listPlanes As New List(Of ePlan)
        'vlSqlParam.Clear()
        'Tabla.Clear()

        'tipoConexion = 1
        'query = "dbo.pa_ListaPlanAdicionales"
        'vlSqlParam.Add("@idTipoContrato", _idTipoContrato.idTipoContrato, SqlDbType.TinyInt)
        'vlSqlParam.Add("@calCodigo", _calCodigo, SqlDbType.Int)

        'Tabla = con.TraeDatosConP(vlSqlParam, query, tipoConexion)
        'For x As Int16 = 0 To Tabla.Rows.Count - 1
        '    Dim row As New ePlan

        '    row.idPlan = IIf(IsDBNull(Tabla.Rows(x)("idPlan")), Nothing, Tabla.Rows(x)("idPlan"))
        '    row.primaUF = IIf(IsDBNull(Tabla.Rows(x)("primaUF")), Nothing, Tabla.Rows(x)("primaUF"))
        '    row.PrimaPesos = IIf(IsDBNull(Tabla.Rows(x)("PrimaPesos")), Nothing, Tabla.Rows(x)("PrimaPesos"))
        '    'row.idTipoContrato = IIf(IsDBNull(dt.Rows(x)("idTipoContrato")), Nothing, dt.Rows(x)("idTipoContrato"))
        '    'row.calCodigo = IIf(IsDBNull(dt.Rows(x)("calCodigo")), Nothing, dt.Rows(x)("calCodigo"))
        '    'row.activo = IIf(IsDBNull(dt.Rows(x)("activo")), Nothing, dt.Rows(x)("activo"))
        '    listPlanes.Add(row)
        'Next

        'Return listPlanes

    End Function
    '***************METODO PARA LISTAR  LOS PLANES DE BENEFICIARIOS********************************
    Public Function ListaPlanBeneficiario(ByVal _idTipoContrato As eTipoContrato, ByVal _calCodigo As Int32) As List(Of ePlan)
        'Dim listPlanes As New List(Of ePlan)
        'vlSqlParam.Clear()
        'Tabla.Clear()

        'tipoConexion = 1
        'query = "dbo.pa_ListaPlanBeneficiario"
        'vlSqlParam.Add("@idTipoContrato", _idTipoContrato.idTipoContrato, SqlDbType.TinyInt)
        'vlSqlParam.Add("@calCodigo", _calCodigo, SqlDbType.Int)

        'Tabla = con.TraeDatosConP(vlSqlParam, query, tipoConexion)

        'For x As Int16 = 0 To Tabla.Rows.Count - 1
        '    Dim row As New ePlan

        '    row.idPlan = IIf(IsDBNull(Tabla.Rows(x)("idPlan")), Nothing, Tabla.Rows(x)("idPlan"))
        '    row.primaUF = IIf(IsDBNull(Tabla.Rows(x)("primaUF")), Nothing, Tabla.Rows(x)("primaUF"))
        '    row.PrimaPesos = IIf(IsDBNull(Tabla.Rows(x)("PrimaPesos")), Nothing, Tabla.Rows(x)("PrimaPesos"))
        '    'row.idTipoContrato = IIf(IsDBNull(Tabla.Rows(x)("idTipoContrato")), Nothing, dt.Rows(x)("idTipoContrato"))
        '    'row.calCodigo = IIf(IsDBNull(dt.Rows(x)("calCodigo")), Nothing, dt.Rows(x)("calCodigo"))
        '    'row.activo = IIf(IsDBNull(dt.Rows(x)("activo")), Nothing, dt.Rows(x)("activo"))
        '    listPlanes.Add(row)
        'Next


        'Return listPlanes

    End Function

End Class
