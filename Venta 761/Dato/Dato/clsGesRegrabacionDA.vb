Imports System.Data.SqlClient
Imports Entidad

Public Class clsGesRegrabacionDA
    Dim con As New clsConexion 'SE CREA OBJETO DE CLSCONEXION PARA PODER CONECTARNOS A BASE DE DATOS
    Dim query As String 'VARIABLE QUE ALMACENARA EL PROCEDIMIENTO ALMACENDADO A CONSULTAR
    Dim tipoConexion As Integer 'VARIABLE QUE ALMACENARA LA BASE A LA CUAL CONSULTAR
    'VARIABLE QUE ALMACENARA LOS PARAMETROS A LA BASE DE DATOS
    Dim dt As New DataTable
    '***********METODO PARA TRAER TODOS LOS DATOS DE LAS VENTAS A SER GESTIONADAS EN REGRABACIONES***************************
    ''' <summary>
    ''' METODO PARA TRAER TODOS LOS DATOS DE LAS VENTAS A SER GESTIONADAS EN REGRABACIONES
    ''' </summary>
    ''' <returns>Datatable</returns>
    ''' <remarks></remarks>
    Public Function Regrabaciones(ByVal usuario As String) As DataTable

        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable


        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Calidad].[pa_BuscaVentasRechazadas]"
        cmd.CommandText = sql
        cmd.Parameters.AddWithValue("@usuario", usuario)
        dt = con.TraeDatos(cmd, 2)

        Return dt

    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="_cli"></param>
    ''' <param name="_rec"></param>
    ''' <param name="_ges"></param>
    Public Sub GuardaClienteGes(ByVal _cli As eCliente, ByVal _rec As eRecuperacion, ByVal _ges As eGestion)

        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable
        Dim valido As Boolean


        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Gestion].[pa_GuardaGestionRecuperacion]"

        cmd.Parameters.AddWithValue("@C_ID", _cli.C_Id)
        cmd.Parameters.AddWithValue("@R_ESTADO", _rec.R_Estado)
        cmd.Parameters.AddWithValue("@R_FECHA", _rec.R_Fecha)
        cmd.Parameters.AddWithValue("@R_HORA", _rec.R_Hora)
        cmd.Parameters.AddWithValue("@R_FECHA_VTA", _rec.R_Fecha_Vta)
        cmd.Parameters.AddWithValue("@R_HORA_VTA", _rec.R_Hora_Vta)
        cmd.Parameters.AddWithValue("@R_IDLLAMADA", _rec.R_IdLlamada)
        cmd.Parameters.AddWithValue("@R_INTENTO", _rec.R_Intento)
        cmd.Parameters.AddWithValue("@R_EJECUTIVO", _rec.R_Ejecutivo)
        cmd.Parameters.AddWithValue("@R_FECHA_AGEN", _rec.R_Fecha_Agen)
        cmd.Parameters.AddWithValue("@R_HORA_AGEN", _rec.R_Hora_Agen)
        cmd.Parameters.AddWithValue("@R_OBS_AGEN", _rec.R_Obs_Agen)
        cmd.Parameters.AddWithValue("@R_CONTACTO", _ges.G_Contacto)
        cmd.Parameters.AddWithValue("@R_NO_CONTACTO", _ges.G_No_Contacto)
        cmd.Parameters.AddWithValue("@R_CONTACTO_CON", _ges.G_Contacto_Con)
        cmd.Parameters.AddWithValue("@R_CONTACTO_FLUJO", _ges.G_Contacto_Flujo)
        cmd.Parameters.AddWithValue("@R_MOTIVO_NO_INTERESA", _ges.G_Motivo_No_Interesa)
        cmd.Parameters.AddWithValue("@R_RECONOCE", _rec.R_Reconoce)
        cmd.Parameters.AddWithValue("@R_VENTA", _rec.R_Venta)
        cmd.CommandText = sql
        valido = con.Ejecutar(cmd, 2)

        If valido = True Then
            MsgBox("Gestion Recuperacion ingresada Correctamente", vbInformation, "CORDIALPHONE")
        Else
            MsgBox("Gestion Recuperacion no fue ingresada, revisar con Supervisor", vbObjectError, "CORDIALPHONE")
        End If

    End Sub
    '************PROCEDIMIENTO PARA ACTUALIZAR LOS CAMPOS EN LA TABLA CLIENTE SIN VENTA  ************************
    ''' <summary>
    ''' PROCEDIMIENTO PARA ACTUALIZAR LOS CAMPOS EN LA TABLA CLIENTE SIN VENTA
    ''' </summary>
    ''' <param name="_seg_est"></param>
    ''' <param name="_est_obj"></param>
    ''' <param name="_call_id"></param>
    ''' <param name="_cli_id"></param>
    ''' <remarks></remarks>
    Public Sub ActualizaCteSinVta(ByVal _seg_est As String, ByVal _est_obj As String, ByVal _call_id As String, ByVal _cli_id As String)

        'vlSqlParam.Clear()
        'query = "[GestionReGrabacion].[ActualizaCteSinVta]"
        'vlSqlParam.Add("@CLI_ID", _cli_id, SqlDbType.BigInt)
        'vlSqlParam.Add("@CLI_SEGUNDO_ESTADO_CALIDAD", _seg_est, SqlDbType.VarChar)
        'vlSqlParam.Add("@CLI_CALL_ID_CALIDAD", _call_id, SqlDbType.VarChar)
        'vlSqlParam.Add("@CLI_ESTADO_OBJECION_CALIDAD", _est_obj, SqlDbType.Int)

        'tipoConexion = 2
        'con.Ejecutar(query, vlSqlParam, tipoConexion)

        'vlSqlParam.Clear()

    End Sub
    '************PROCEDIMIENTO PARA ACTUALIZAR LOS CAMPOS EN LA TABLA AsignacionCalidad  ************************
    Public Sub GrabaAsignaCalidad(ByVal _cliid As String, ByVal _calcodigo_centro As Integer, ByVal _agent As String, ByVal fecha As String, ByVal opcion As Integer, ByVal _agent_vta As String)

        'vlSqlParam.Clear()
        'query = "[dbo].[pa_Inserta_AsigCalidad_Rechaza]"
        'vlSqlParam.Add("@cliid", _cliid, SqlDbType.BigInt)
        'vlSqlParam.Add("@calCodigo", _calcodigo_centro, SqlDbType.Int)
        'vlSqlParam.Add("@agente", _agent, SqlDbType.VarChar)
        'vlSqlParam.Add("@fechavta", fecha, SqlDbType.VarChar)
        'vlSqlParam.Add("@Opcion", opcion, SqlDbType.Int)
        'vlSqlParam.Add("@Agente_Venta", _agent_vta, SqlDbType.VarChar)

        'tipoConexion = 1
        'con.Ejecutar(query, vlSqlParam, tipoConexion)

        'vlSqlParam.Clear()

    End Sub
    '*********************METODO PARA ACTUALIZAR CLIENTE AGENDADO DE REGRABACIONES***********
    Public Sub ActualizaClienteAgen(ByVal _cliid As String, ByVal _est_obj As String, ByVal _call_id As String)

        'vlSqlParam.Clear()
        'query = "[GestionReGrabacion].[ActualizaClienteAgen]"
        'vlSqlParam.Add("@cli_id ", _cliid, SqlDbType.BigInt)
        'vlSqlParam.Add("@CLI_ESTADO_OBJECION_CALIDAD", _est_obj, SqlDbType.Int)
        'vlSqlParam.Add("@CLI_CALL_ID_CALIDAD", _call_id, SqlDbType.VarChar)

        'tipoConexion = 2
        'con.Ejecutar(query, vlSqlParam, tipoConexion)

        'vlSqlParam.Clear()

    End Sub
    '*********************METODO PARA ACTUALIZAR CLIENTE ACON VENTA EN REGRABACIONES***********
    Public Sub ActualizaClienteVenta(ByVal _cliente As eCliente, ByVal _gestion As eGestion)

        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable
        Dim valido As Boolean


        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Gestion].[pa_ActualizaClienteRecuperacion]"

        cmd.Parameters.AddWithValue("@C_ID", _cliente.C_Id)
        cmd.Parameters.AddWithValue("@G_RUT", _gestion.G_Rut)
        cmd.Parameters.AddWithValue("@G_DV", _gestion.G_Dv)
        cmd.Parameters.AddWithValue("@G_PATERNO", _gestion.G_Paterno)
        cmd.Parameters.AddWithValue("@G_MATERNO", _gestion.G_Materno)
        cmd.Parameters.AddWithValue("@G_NOMBRE", _gestion.G_Nombre)
        cmd.Parameters.AddWithValue("@G_FECHANACIMIENTO", _gestion.G_Fecha_Nacimiento)
        cmd.Parameters.AddWithValue("@G_SEXO", _gestion.G_Sexo)
        cmd.Parameters.AddWithValue("@G_EMAIL", _gestion.G_Email)
        cmd.Parameters.AddWithValue("@G_CALLE", _gestion.G_Calle)
        cmd.Parameters.AddWithValue("@G_NRO", _gestion.G_Nro)
        cmd.Parameters.AddWithValue("@G_REFERENCIA", _gestion.G_Referencia)
        cmd.Parameters.AddWithValue("@G_COMUNA", _gestion.G_Comuna)
        cmd.Parameters.AddWithValue("@G_CIUDAD", _gestion.G_Ciudad)
        cmd.Parameters.AddWithValue("@G_IDCOMUNA", _gestion.G_IdComuna)
        cmd.Parameters.AddWithValue("@G_IDCIUDAD", _gestion.G_IdCiudad)
        cmd.Parameters.AddWithValue("@G_FONO_CONTACTO", _gestion.G_Fono_Contacto)
        cmd.Parameters.AddWithValue("@G_DATO1", _gestion.G_Dato1)
        cmd.Parameters.AddWithValue("@G_DATO2", _gestion.G_Dato2)
        cmd.Parameters.AddWithValue("@G_DATO3", _gestion.G_Dato3)
        cmd.Parameters.AddWithValue("@G_BANCO", _gestion.G_BANCO)
        cmd.Parameters.AddWithValue("@G_VENCIMIENTO_TARJETA", _gestion.G_VENCIMIENTO_TARJETA)
        cmd.Parameters.AddWithValue("@G_NUMERO_TARJETA", _gestion.G_NUMERO_TARJETA)
        cmd.Parameters.AddWithValue("@G_MEDIO_PAGO", _gestion.G_MEDIO_PAGO)
        cmd.Parameters.AddWithValue("@G_TIPO_TARJETA", _gestion.G_TIPO_TARJETA)
        cmd.Parameters.AddWithValue("@G_DIA_PAGO", _gestion.G_DIA_PAGO)

        cmd.CommandText = sql
        valido = con.Ejecutar(cmd, 2)

        If valido = True Then
            MsgBox("Datos CLiente Actualizados ", vbInformation, "CORDIALPHONE")
        Else
            MsgBox("Datos no fueron actualizados, revisar con Supervisor", vbObjectError, "CORDIALPHONE")
        End If

    End Sub

End Class
