Imports System.Data.SqlClient
Imports Entidad

Public Class clsGeneralDA
    Dim con As New clsConexion 'SE CREA OBJETO DE CLSCONEXION PARA PODER CONECTARNOS A BASE DE DATOS
    Dim query As String 'VARIABLE QUE ALMACENARA EL PROCEDIMIENTO ALMACENDADO A CONSULTAR
    Dim tipoConexion As Integer 'VARIABLE QUE ALMACENARA LA BASE A LA CUAL CONSULTAR
    'VARIABLE QUE ALMACENARA LOS PARAMETROS A LA BASE DE DATOS
    Dim Tabla As New DataTable
    '*********METODO PARA OBTENER LA FECHA DEL SERVIDOR*******************
    Function FechaServidor() As DateTime
        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable
        Dim fecha As DateTime

        cmd.CommandType = CommandType.StoredProcedure
        sql = "[Gestion].[Fecha_Servidor]"
        cmd.CommandText = sql
        dt = con.TraeDatos(cmd, 1)

        fecha = IIf(IsDBNull(dt.Rows(0)("Fecha")), Nothing, dt.Rows(0)("Fecha"))

        Return fecha
    End Function
    '*********METODO PARA GUARDAR LLAMADA EN LA TABLA CID*******************
    Public Sub Insertar_Grabacion(ByVal _grabacion As eLlamada, ByVal tipo As Integer)


        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable
        Dim listTipoContrato As New List(Of eTipoContrato)


        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Gestion].[pa_RegistraLLamada]"
        cmd.Parameters.AddWithValue("@tipo", tipo)
        cmd.Parameters.AddWithValue("@C_ID", _grabacion.C_ID)
        cmd.Parameters.AddWithValue("@LL_IDLLAMADA", _grabacion.LL_CallId)
        cmd.Parameters.AddWithValue("@LL_FECHA", _grabacion.LL_Fecha)
        cmd.Parameters.AddWithValue("@LL_EJECUTIVO", _grabacion.LL_Ejecutivo)
        cmd.Parameters.AddWithValue("@LL_HORA_INICIO", _grabacion.LL_Hora_Inicio)
        cmd.Parameters.AddWithValue("@LL_HORA_TERMINO", _grabacion.LL_Hora_Termino)
        cmd.Parameters.AddWithValue("@LL_FONO_CONTACTO", _grabacion.LL_Fono_Contacto)
        cmd.CommandText = sql
        con.Ejecutar(cmd, 2)

    End Sub
    '*********METODO PARA GUARDAR LLAMADA EN LA TABLA LOGUEO_USUARIO*******************
    Public Sub Logear_Usuario(ByRef Usuario As String, ByRef estado As Short)

        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable

        tipoConexion = 2

        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Sesion].[sp_LogeoUsuario]"
        cmd.Parameters.AddWithValue("@USUARIO", Usuario)
        cmd.Parameters.AddWithValue("@ESTLOG", estado)
        cmd.CommandText = sql
        con.Ejecutar(cmd, 2)
    End Sub
    '*********METODO PARA TRAER LOS PARENTESCOS*******************
    Function Listar_Parentesco() As DataTable
        'Tabla.Clear()
        'tipoConexion = 1
        '    query = "[dbo].[pa_ListarParentesco]"
        'Tabla = con.TraeDatosSinP(query, tipoConexion)
        '    Return Tabla
    End Function
    '*********METODO PARA TRAER LOS PARENTESCOS PARA LOS BENEFICIARIOS*******************
    Function Listar_ParentescoBen() As DataTable

        'Tabla.Clear()

        'tipoConexion = 1
        'query = "[dbo].[pa_ListarParentescoBen]"
        'Tabla = con.TraeDatosSinP(query, tipoConexion)

        'Return Tabla
    End Function

    '*********METODO PARA ACTUALIZAR ESTADO DE CLI_ESTADO DE LA TABLA CLI EN CASO DE QUE SE PRESIONE BOTON SALIR*******************
    Public Sub Respaldar_Estado(ByVal estado As String, ByVal id As UInteger)
        'vlSqlParam.Clear()

        'tipoConexion = 2
        'vlSqlParam.Add("@estado", estado, SqlDbType.VarChar)
        'vlSqlParam.Add("id", id, SqlDbType.BigInt)
        'query = "[Config].[Respaldar_estado_anterior]"
        'con.Ejecutar(query, vlSqlParam, tipoConexion)
        'vlSqlParam.Clear()
    End Sub
    '*********METODO PARA BUSCAR VALOR UF*******************
    Public Function Busca_ValorUF() As String
        'Dim s As String
        's = ""
        'vlSqlParam.Clear()
        'Tabla.Clear()
        'query = "dbo.pa_Busca_ValorUF "
        'tipoConexion = 1
        'Tabla = con.TraeDatosSinP(query, tipoConexion)

        'If Tabla.Rows.Count > 0 Then
        '    s = Tabla.Rows(0)("UF")
        'End If

        'Return s
    End Function

    Public Sub IngresarDescanso(ByRef Usuario As String, ByRef tipoDescanso As Integer, ByVal descanso As Boolean)

        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable

        tipoConexion = 2

        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Sesion].[sp_IngresoDescanso]"
        cmd.Parameters.AddWithValue("@Usuario", Usuario)
        cmd.Parameters.AddWithValue("@tipoDescanso", tipoDescanso)
        cmd.Parameters.AddWithValue("@descanso", descanso)
        cmd.CommandText = sql
        con.Ejecutar(cmd, 2)

    End Sub

End Class
