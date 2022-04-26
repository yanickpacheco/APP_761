Imports System.Data.SqlClient
Imports Entidad

Public Class clsClienteDA
    Dim con As New clsConexion
    Dim query As String
    Dim tipoConexion As Integer

    '********* Guarda los datos obtenidos de la venta ***************
    ''' <summary>
    ''' GUARDA LOS DATOS OBTENIDOS DE LA APLICACION DE VENTAS EN LA TABLA CLI
    ''' </summary>
    ''' <param name="_cliente"></param>
    ''' <remarks></remarks>
    Public Sub GuardaDatosCliente(ByVal _cliente As eCliente, ByVal _gestion As eGestion)


        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable
        Dim valido As Boolean


        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Gestion].[pa_GuardaGestion]"

        cmd.Parameters.AddWithValue("@C_ID", _cliente.C_Id)
        cmd.Parameters.AddWithValue("@G_ESTADO", _gestion.G_Estado)
        cmd.Parameters.AddWithValue("@G_FECHA", _gestion.G_Fecha)
        cmd.Parameters.AddWithValue("@G_HORA", _gestion.G_Hora)
        cmd.Parameters.AddWithValue("@G_FECHA_VTA", _gestion.G_Fecha_Vta)
        cmd.Parameters.AddWithValue("@G_HORA_VTA", _gestion.G_Hora_Vta)
        cmd.Parameters.AddWithValue("@G_CALLID", _gestion.G_IdLlamada)
        cmd.Parameters.AddWithValue("@G_INTENTO", _gestion.G_Intento)
        cmd.Parameters.AddWithValue("@G_EJECUTIVO", _gestion.G_Ejecutivo)
        cmd.Parameters.AddWithValue("@G_IP_EJECUTIVO", _gestion.G_Ip_Ejecutivo)
        cmd.Parameters.AddWithValue("@G_FECHA_AGEN", _gestion.G_Fecha_Agen)
        cmd.Parameters.AddWithValue("@G_HORA_AGEN", _gestion.G_Hora_Agen)
        cmd.Parameters.AddWithValue("@G_OBS_AGEN", _gestion.G_Obs_Agen)
        cmd.Parameters.AddWithValue("@G_CONTACTO", _gestion.G_Contacto)
        cmd.Parameters.AddWithValue("@G_NO_CONTACTO", _gestion.G_No_Contacto)
        cmd.Parameters.AddWithValue("@G_CONTACTO_CON", _gestion.G_Contacto_Con)
        cmd.Parameters.AddWithValue("@G_CONTACTO_FLUJO", _gestion.G_Contacto_Flujo)
        cmd.Parameters.AddWithValue("@G_MOTIVO_NO_INTERESA", _gestion.G_Motivo_No_Interesa)
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
        cmd.Parameters.AddWithValue("@G_IDCOMUNA", _gestion.G_IdComuna)
        cmd.Parameters.AddWithValue("@G_IDCIUDAD", _gestion.G_IdCiudad)
        cmd.Parameters.AddWithValue("@G_CIUDAD", _gestion.G_Ciudad)
        cmd.Parameters.AddWithValue("@G_FONO_CONTACTO", _gestion.G_Fono_Contacto)
        cmd.Parameters.AddWithValue("@G_TIPOCONTRATO", _gestion.G_TipoContrato)
        cmd.Parameters.AddWithValue("@G_PLAN", _gestion.G_Plan)
        cmd.Parameters.AddWithValue("@G_PRIMA_UF", _gestion.G_Prima_Uf)
        cmd.Parameters.AddWithValue("@G_PRIMA_PESOS", _gestion.G_Prima_Pesos)
        cmd.Parameters.AddWithValue("@G_VENTA", _gestion.G_Venta)
        cmd.Parameters.AddWithValue("@G_DATO1", _gestion.G_Dato1)
        cmd.Parameters.AddWithValue("@G_DATO2", _gestion.G_Dato2)
        cmd.Parameters.AddWithValue("@G_DATO3", _gestion.G_Dato3)
        cmd.Parameters.AddWithValue("@G_BANCO", _gestion.G_BANCO)
        cmd.Parameters.AddWithValue("@G_VENCIMIENTO_TARJETA", _gestion.G_VENCIMIENTO_TARJETA)
        cmd.Parameters.AddWithValue("@G_NUMERO_TARJETA", _gestion.G_NUMERO_TARJETA)
        cmd.Parameters.AddWithValue("@G_MEDIO_PAGO", _gestion.G_MEDIO_PAGO)
        cmd.Parameters.AddWithValue("@G_TIPO_TARJETA", _gestion.G_TIPO_TARJETA)
        cmd.Parameters.AddWithValue("@G_DIA_PAGO", _gestion.G_DIA_PAGO)
        cmd.Parameters.AddWithValue("@G_TIEMPO_POSTVIEW", _gestion.G_Tiempo_PostView)
        cmd.CommandText = sql
        valido = con.Ejecutar(cmd, 2)

        If valido Then
            MsgBox("Gestion ingresada Correctamente", vbInformation, "CORDIALPHONE")
        Else
            MsgBox("Gestion no fue ingresada, revisar con Supervisor", vbObjectError, "CORDIALPHONE")
        End If

    End Sub

    '********* Guarda los datos obtenidos de la venta en la tabla log ***************
    Public Sub GuardaDatosLog(ByVal _claveRegistroActual As String)
        'vlSqlParam.Clear()
        'query = "[GestionVenta].SP_" + nomcampania + "_LOG"
        'vlSqlParam.Add("@id", _claveRegistroActual, SqlDbType.Int)
        'vlSqlParam.Add("@fecha_ter_gestion", DateTime.Now.ToString("yyyyMMdd"), SqlDbType.VarChar)
        'vlSqlParam.Add("@hora_ter_gestion", DateTime.Now.ToString("HHmmss"), SqlDbType.VarChar)
        'tipoConexion = 2
        'con.Ejecutar(query, vlSqlParam, tipoConexion)
        'vlSqlParam.Clear()

    End Sub
    '********* Obtiene los datos del cliente para la venta***************
    Public Function BuscarCliente(ByVal WS_IDUSUARIO As String) As DataTable
        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable


        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Gestion].[pa_BuscaCliente]"
        cmd.Parameters.AddWithValue("@usuario", WS_IDUSUARIO)
        cmd.CommandText = sql
        dt = con.TraeDatos(cmd, 2)

        If dt.Rows.Count <= 0 Then
            MsgBox("No hay registros por recorrer", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "CORDIAL PHONE")
            Return dt
        ElseIf dt.Rows(0)("C_ID") = 0 Then
            MsgBox("No hay registros por recorrer", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "CORDIAL PHONE")
            Return dt
        End If
        Return dt
    End Function

    '*********METODO PARA OBTENER CLIENTE PARA REGRABACION*******************
    Public Function GesCliente(ByVal Usuario As String, ByVal id As Integer) As DataTable
        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable


        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Gestion].[pa_BuscaClienteRecuperacion]"
        cmd.Parameters.AddWithValue("@USUARIO", Usuario)
        cmd.Parameters.AddWithValue("@id", id)
        cmd.CommandText = sql
        dt = con.TraeDatos(cmd, 2)

        Return dt

        'vlSqlParam.Clear()
        'query = "[GestionReGrabacion].[Traer_cliente_VentaR]"
        'vlSqlParam.Add("@id", id, SqlDbType.Int)
        'tipoConexion = 2
        'dt = con.TraeDatosConP(vlSqlParam, query, tipoConexion)
        'xt.G_ID = IIf(IsDBNull(dt.Rows(0)("CLI_ID")), Nothing, dt.Rows(0)("CLI_ID"))
        'xt.G_Nombre = IIf(IsDBNull(dt.Rows(0)("CLI_ANOMBRE")), Nothing, dt.Rows(0)("CLI_ANOMBRE"))
        'xt.G_Paterno = IIf(IsDBNull(dt.Rows(0)("CLI_APATERNO")), Nothing, dt.Rows(0)("CLI_APATERNO"))
        'xt.G_Materno = IIf(IsDBNull(dt.Rows(0)("CLI_AMATERNO")), Nothing, dt.Rows(0)("CLI_AMATERNO"))
        'xt.G_Rut = IIf(IsDBNull(dt.Rows(0)("CLI_ARUT")), Nothing, dt.Rows(0)("CLI_ARUT"))
        'xt.G_Dv = IIf(IsDBNull(dt.Rows(0)("CLI_ADV")), Nothing, dt.Rows(0)("CLI_ADV"))
        'xt.G_Comuna = IIf(IsDBNull(dt.Rows(0)("CLI_ACOMUNA")), Nothing, dt.Rows(0)("CLI_ACOMUNA"))
        'xt.G_Ciudad = IIf(IsDBNull(dt.Rows(0)("CLI_ACIUDAD")), Nothing, dt.Rows(0)("CLI_ACIUDAD"))
        'xt.G_Fecha_Nacimiento = IIf(IsDBNull(dt.Rows(0)("CLI_AFECHANACIMIENTO")), Nothing, dt.Rows(0)("CLI_AFECHANACIMIENTO"))
        'xt.G_Sexo = IIf(IsDBNull(dt.Rows(0)("CLI_ASEXO")), Nothing, dt.Rows(0)("CLI_ASEXO"))
        ''xt.G_codverificacion = IIf(IsDBNull(dt.Rows(0)("CLI_CODVERIFICACION")), Nothing, dt.Rows(0)("CLI_CODVERIFICACION"))
        'xt.G_Calle = IIf(IsDBNull(dt.Rows(0)("CLI_ACALLE")), Nothing, dt.Rows(0)("CLI_ACALLE"))
        'xt.G_Nro = IIf(IsDBNull(dt.Rows(0)("CLI_ANRO")), Nothing, dt.Rows(0)("CLI_ANRO"))
        'xt.G_Referencia = IIf(IsDBNull(dt.Rows(0)("CLI_AREFERENCIA")), Nothing, dt.Rows(0)("CLI_AREFERENCIA"))
        ''xtGI_ATARJETAVENCIMIENTO = IIf(IsDBNull(dt.Rows(0)("CLI_ATARJETAVENCIMIENTO")), Nothing, dt.Rows(0)("CLI_ATARJETAVENCIMIENTO"))
        'xt.G_Prima_Pesos = IIf(IsDBNull(dt.Rows(0)("CLI_PRIMAPESOS")), Nothing, dt.Rows(0)("CLI_PRIMAPESOS"))
        'xt.G_Prima_Uf = IIf(IsDBNull(dt.Rows(0)("CLI_PRIMAUF")), Nothing, dt.Rows(0)("CLI_PRIMAUF"))
        'xt.G_Plan = IIf(IsDBNull(dt.Rows(0)("CLI_PLAN")), Nothing, dt.Rows(0)("CLI_PLAN"))
        '' xt.G_tpocontrato = IIf(IsDBNull(dt.Rows(0)("CLI_TPOCONTRATO")), Nothing, dt.Rows(0)("CLI_TPOCONTRATO"))
        'xt.G_Email = IIf(IsDBNull(dt.Rows(0)("CLI_AEMAIL")), Nothing, dt.Rows(0)("CLI_AEMAIL"))
        ''xtGcli_agen_obs = IIf(IsDBNull(dt.Rows(0)("CLI_OBSERVACION")), Nothing, dt.Rows(0)("CLI_OBSERVACION"))
        'xt.G_Fono_Contacto = IIf(IsDBNull(dt.Rows(0)("CLI_AFONOCONTACTO")), Nothing, dt.Rows(0)("CLI_AFONOCONTACTO"))


        'Return xt
    End Function

End Class
