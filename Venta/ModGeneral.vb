Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.Collections.Generic
Imports Entidad
Imports BI

Module ModGeneral

    Public db_central As Short = My.Settings.vgCentral
    Public vgCampania As eProyecto
    Public vgListPlanes As New List(Of ePlan)
    Public vgListTipoContrato As New List(Of eTipoContrato)
    Public vgListTipoBeneficiario As New List(Of eTipoParentesco)
    Public vgConfiguracionCampania As New clsConfiguracionCampaniaBI
    Public usCodigo As String
    Dim biGeneral As New clsGeneralBI
    Public nombreBase As String
    Public planE As New ePlan
    Public vgFuncionComun As New clsFuncionesComunes
    Private nroarchivo As Short
    Public NroAseg As Short
    Public totalUfAdic As Double = 0
    Public ePlanGlobal As New ePlan
    Public valorPesosUf As Integer = 0
    Public TotalPesos As Integer
    Public tipoContrato As New eTipoContrato
    Public ufAdic As Double = 0
    Public PesosAdic As Int64
    Public idPlanAdic As Int64

    'Para guardar informacion de los Beneficiarios del titular
    Structure DatosB
        Dim bnro As String
        Dim bnombre1 As String
        Dim bnombre2 As String
        Dim bappaterno As String
        Dim bapmaterno As String
        Dim brut As String
        Dim bdv As String
        Dim bpctje As String
        Dim bparentesco As String
        Dim bparentxt As String
        Dim biparentesco As String
        Dim bOtroparentesco As String
    End Structure

    Public Structure tUsuario
        Dim Computer As String
        Dim UserName As String
        Dim Fullname As String
        Dim Domain As String
        Dim PDC As String
        Dim IP As String
    End Structure

    Public Const csNombreAplicacion As String = "SEGURO ONCOLOGICO CLINICA LAS CONDES"

    'Para los valores de la Uf
    Public vlFechaUd As String

    Public Const csEdadMaxima As Short = 80
    Public Const csEdadMinima As Short = 18

    Public Const csEdadMinimaAdi As Short = 0
    Public Const csEdadMaximaAdi As Short = 79

    Public lstAdi As New List(Of eAdicional)
    Public lstBen As New List(Of eBeneficiario)
    Public listParentescoCampania As List(Of eTipoParentesco)
    Public FilaAgregar As Integer

    Public numeroFila As Integer

    Public perfil As String 'VARIABLE QUE ALMACENARA EL PERFIL DEL AGENTE
    Public GesId As Integer 'VARIABLE QUE ALMACENARA LA ID A EVALUAR EN REGRABACIONES
    Public estado_perfil As Boolean 'VARIABLE QUE INDICARA SI SE ABRE FORMULARIO DE VENTAS O DE REGRABACIONES
    Public preex As String = "" 'variable que almacenara un valor para indicar si cliente tiene preexistencia esto es para regrabaciones

    Public FecActual As Date

    Public Const csValSoloAlfa As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZÑñ "
    Public Const csValSimbolos As String = ".,º-_()@"
    Public Const csValAlfaNumericos As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789Ññ .,º-_()@"
    Public Const csValNumeros As String = "0123456789"
    Public Const csValDigRut As String = "0123456789Kk"

    Public flag_primeravez As Boolean
    Public flag_virgen As Boolean
    Public flg_progresivo_activado As Boolean
    Public flg_fono1 As Boolean
    Public flg_EsCelu1 As Boolean
    Public flg_fono2 As Boolean
    Public flg_EsCelu2 As Boolean
    Public flg_fono3 As Boolean
    Public flg_EsCelu3 As Boolean
    Public flg_fono4 As Boolean
    Public flg_EsCelu4 As Boolean
    Public flg_fono5 As Boolean
    Public flg_EsCelu5 As Boolean
    Public flg_fono6 As Boolean
    Public flg_EsCelu6 As Boolean
    Public flg_fonoalt As Boolean
    Public flg_EsCeluAlt As Boolean
    Public flg_fonoVent As Boolean
    Public flg_EsCeluVent As Boolean

    Public sql As String

    Public Const CENTRODECOSTO As String = "761"

    Public claveRegistroActual As String

    Public WS_IP As String
    Public WS_APLICACION As String
    Public WS_IDAPLICACION As String
    Public WS_IDUSUARIO As String
    Public WS_NOMUSUARIO As String
    Public WS_TIPO_APLICACION As String
    Public WS_FONO_LLAMA As String
    Public WS_CALL_ID As String
    Public WS_CAMPAIGN_CIC As String
    Public WS_CicServer As String
    Public WS_Usuario As String
    Public WS_Pass As String
    Public WS_Maquina As String
    Public Datos As String
    Public Fono_A_Llamar As String
    Public EdadA As String

    ' guarda string sql update para restaurart estado inicial
    Public strQueryUpdateBackupRs As String

    ' implementacion de estructura de pila
    Public pila() As Integer
    Public pila_cuenta As Integer

    Public ArrTarjetaDeCredito() As String ' Arreglo con las tarjetas de creditos
    Public usuario_actual As tUsuario

    Public oSoapClient As New WSSNeoTel.WebService
    Public vpPosicion As New clsSoapNeo
    Public vpNeotel As New clsSoapNeo
    Public vgUsuario As Short

    Public vgListCiudad As New List(Of eCiudad)
    Public vgListComuna As New List(Of eComuna)
    Public vgListEdoFono As New List(Of eEstadoFonos)
    Public vgListParentescoCampania As New List(Of eTipoParentesco)

    Public feriados() As Date
    Public idDescanso As Integer


    ' #####################################################################################
    '
    ' #####################################################################################

    ' ----------------------------------
    Public Sub Main()
        configuraUsuarioActual()
        vgCampania = vgConfiguracionCampania.ObtenerConfiguracionCampania(My.Settings.vgCRM)
        oSoapClient.Url = "http://190.98.236.19/neoapi/webservice.asmx?wsdl"

        'Establece_conexion
        frmLogin.ShowDialog()
        System.Windows.Forms.Application.DoEvents()
    End Sub



    'obtenemos informacion de la maquina o usuario
    'a traves del dominio

    Private Sub configuraUsuarioActual()

        With usuario_actual
            'My.User.Name
            'My.Computer.Name 
            'My.Application.Info.

            .Computer = UCase(My.Computer.Name)
            .UserName = UCase(Environment.UserName)
            .Domain = UCase(Environment.UserDomainName)
            .PDC = UCase(My.Computer.Name) 'UCase(sysInfo.PDC)
        End With
        WS_IDUSUARIO = usuario_actual.Computer
        usuario_actual.IP = vgFuncionComun.getIp
        usuario_actual.Fullname = usuario_actual.UserName

        ' WS_IDUSUARIO = "YANICK.PACHECO"

    End Sub

    Public Sub conectarTelefonia()
        If Not vpPosicion.Login_Campania((vpPosicion.Usuario), vgCampania.IdTelefonia) Then 'Numero de campaña Neo
            End
        End If
    End Sub


    Function FechaServidor() As DateTime
        Return biGeneral.BuscarFecha
    End Function

    Function validaCaracteres(ByRef TxtObj As System.Windows.Forms.TextBox, ByRef constValString As String) As Boolean
        Dim flg_ok As Boolean
        Dim indice As Integer

        Dim chr_ok As String
        chr_ok = constValString

        Dim nueva_cadena As String
        nueva_cadena = ""

        flg_ok = True
        ' parsea la cadena, si encuentra un caracter que no este en el conjunto, no es valida
        For indice = 1 To Len(TxtObj.Text)
            If (InStr(1, chr_ok, Mid(TxtObj.Text, indice, 1)) = 0) Then
                flg_ok = False
            Else
                nueva_cadena = nueva_cadena & Mid(TxtObj.Text, indice, 1)
            End If
        Next indice

        TxtObj.Text = nueva_cadena
        validaCaracteres = flg_ok

    End Function

    'UPGRADE_NOTE: str se actualizó a str_Renamed. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Function esVacio(ByRef str_Renamed As String) As Boolean

        'UPGRADE_WARNING: Se detectó el uso de Null/IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        esVacio = IIf(Len(Trim(IIf(IsDBNull(str_Renamed), "", str_Renamed))) = 0, True, False)

    End Function

    Function esCelular(ByRef str_Renamed As String) As Boolean
        esCelular = IIf(Len(Trim(IIf(IsDBNull(str_Renamed), "", str_Renamed))) > 7 And Mid(str_Renamed, 1, 1) = "9", True, False)
    End Function


    Public Function DesplegarEdad(ByRef sFecha As String) As String
        Dim FechaNacimiento As Date
        Dim fechaActual As Date
        Dim iDiferencia As Short
        Dim iDiferenciaAgnos As Short
        Dim iDiferenciaMeses As Short
        DesplegarEdad = "Edad: "

        If IsDate(sFecha) Then

            FechaNacimiento = CDate(sFecha)
            fechaActual = Now
            iDiferencia = DateDiff(Microsoft.VisualBasic.DateInterval.Month, FechaNacimiento, fechaActual, FirstDayOfWeek.Monday)
            iDiferenciaAgnos = iDiferencia \ 12
            iDiferenciaMeses = iDiferencia Mod 12

            If iDiferenciaMeses = 0 Then
                DesplegarEdad = "Edad: " & Trim(CStr(iDiferenciaAgnos)) & " años."
            Else
                DesplegarEdad = "Edad: " & Trim(CStr(iDiferenciaAgnos)) & " años y " & Trim(CStr(iDiferenciaMeses)) & " meses."
            End If

        End If

    End Function

    Public Function edad(ByVal fchNacim As Date) As Integer
        edad = Year(Today) - Year(fchNacim)
        If Month(Today) < Month(fchNacim) Then
            edad = edad - 1
        ElseIf Month(Today) = Month(fchNacim) And VB.Day(Today) < VB.Day(fchNacim) Then
            edad = edad - 1
        End If
    End Function

    Function ValidEmail(ByVal strCheck As String) As Boolean
        'Created by Chad M. Kovac
        'Tech Knowledgey, Inc.
        'http://www.TechKnowledgeyInc.com

        Dim bCK As Boolean
        Dim strDomainType As String
        Const sInvalidChars As String = "!#$%^&*()=+{}[]|\;:'/?>,< "
        Dim i As Short

        bCK = Not InStr(1, strCheck, Chr(34)) > 0 'Check to see if there is a double quote
        If Not bCK Then GoTo ExitFunction

        bCK = Not InStr(1, strCheck, "..") > 0 'Check to see if there are consecutive dots
        If Not bCK Then GoTo ExitFunction

        ' Check for invalid characters.
        If Len(strCheck) > Len(sInvalidChars) Then
            For i = 1 To Len(sInvalidChars)
                If InStr(strCheck, Mid(sInvalidChars, i, 1)) > 0 Then
                    bCK = False
                    GoTo ExitFunction
                End If
            Next
        Else
            For i = 1 To Len(strCheck)
                If InStr(sInvalidChars, Mid(strCheck, i, 1)) > 0 Then
                    bCK = False
                    GoTo ExitFunction
                End If
            Next
        End If

        If InStr(1, strCheck, "@") > 1 Then 'Check for an @ symbol
            bCK = Len(Left(strCheck, InStr(1, strCheck, "@") - 1)) > 0
        Else
            bCK = False
        End If
        If Not bCK Then GoTo ExitFunction

        strCheck = Right(strCheck, Len(strCheck) - InStr(1, strCheck, "@"))
        bCK = Not InStr(1, strCheck, "@") > 0 'Check to see if there are too many @'s
        If Not bCK Then GoTo ExitFunction

        strDomainType = Right(strCheck, Len(strCheck) - InStr(1, strCheck, "."))
        bCK = Len(strDomainType) > 0 And InStr(1, strCheck, ".") < Len(strCheck)
        If Not bCK Then GoTo ExitFunction

        On Error GoTo Error2
        strCheck = Left(strCheck, Len(strCheck) - Len(strDomainType) - 1)
        On Error GoTo 0
        Do Until InStr(1, strCheck, ".") <= 1
            If Len(strCheck) >= InStr(1, strCheck, ".") Then
                strCheck = Left(strCheck, Len(strCheck) - (InStr(1, strCheck, ".") - 1))
            Else
                bCK = False
                GoTo ExitFunction
            End If
        Loop
        If strCheck = "." Or Len(strCheck) = 0 Then bCK = False
        GoTo ExitFunction

Error2:
        bCK = False
        GoTo ExitFunction

ExitFunction:
        ValidEmail = bCK
    End Function

    Public Function generarQueryBackupRs(ByVal tblBackup As DataTable, ByRef recordId As String) As String
        Dim i As Integer
        Dim ponerComa As Boolean
        Dim estado As String

        Try
            ponerComa = False
            estado = ""
            'sqlUpdate = "update " & TACLIENTE & " set "
            For i = 0 To tblBackup.Rows.Count - 1

                'Select Case UCase(tblBackup.Rows(i)(""))
                '    Case CAMPOCLAVE, "CLI_FECHA_CARGA", "CLI_PRIMERA_FECHA_CALIDAD", "CLI_PRIMERA_FECHA_CALIDAD", "CLI_PRIMER_ESTADO_CALIDAD", "CLI_SEGUNDA_FECHA_CALIDAD", "CLI_SEGUNDO_ESTADO_CALIDAD", "CLI_ESTADO_OBJECION_CALIDAD", "CLI_TELEFONO_ALT", "CLI_FECHA_VENCIMIENTO_DATA", "CLI_VENTA_INFORMADA", "CLI_FECHA_VENTA_INFORMADA", "CLI_BLOQUEADO", "CLI_ANIOMES", "CLI_ORDEN_CAIDA"

                '        ' saltar este campo
                '    Case Else

                'If ponerComa Then sqlUpdate = sqlUpdate & ", "
                'If Not ponerComa Then ponerComa = True

                ' si el campo es la cli_estado y no trae fecha de agendamiento, dejar como disponible
                If tblBackup.Rows(i)("CLI_AGEN_FECHA") & "" = "" Then
                    'If UCase(backup.Fields(i).Name) = "CLI_ESTADO" And ( IsDBNull(backup.Fields("CLI_AGEN_FECHA").Value) = True or  Trim(backup.Fields("CLI_AGEN_FECHA").Value)) = "" ) Then
                    'sqlUpdate = sqlUpdate & "CLI_ESTADO" & " = 'D' "
                    estado = "D"
                Else
                    ' si el campo es ESTADO y SI TRAE fecha de agendamiento dejar como Estado = A                            
                    If tblBackup.Rows(i)("CLI_AGEN_FECHA") & "" <> "" Then
                        'sqlUpdate = sqlUpdate & "CLI_ESTADO" & " = 'A' "
                        estado = "A"
                    End If
                End If

                'End Select
            Next i
            'sqlUpdate = sqlUpdate & " where " & CAMPOCLAVE & " = " & recordId
            'generarQueryBackupRs = sqlUpdate
            generarQueryBackupRs = estado
        Catch ex As Exception
            MsgBox(Err.Description & " " & " Error : En la Función generarQueryBackupRs()", MsgBoxStyle.Critical, "CORDIALPHONE")
            generarQueryBackupRs = ""
        End Try
    End Function



    Public Sub grabarCallId(ByRef accion As String, ByRef callid As String, ByRef fonoContacto As String, ByRef clave As String)
        Dim cid As New eLlamada
        Dim ahora As DateTime
        Dim tipo As UShort
        Dim t As String


        Try
            If callid.Length > 1 Then
                t = ""
                tipo = 0
                ahora = FechaServidor()
                cid.C_ID = clave
                cid.LL_CallId = callid
                cid.LL_Fono_Contacto = fonoContacto
                cid.LL_Fecha = ahora.ToString("yyyyMMdd")
                cid.LL_Hora_Inicio = ahora.ToString("HHmmss")
                cid.LL_Ejecutivo = WS_IDUSUARIO
                Select Case accion
                    Case "LLAMAR"
                        tipo = 1
                        t = ""
                    Case "CORTAR"
                        tipo = 2
                        t = ahora.ToString("HHmmss")
                End Select
                cid.LL_Hora_Termino = t
                biGeneral.InsertarGrabacion(cid, tipo)
            End If
        Catch ex As Exception
            MsgBox(Err.Description & " " & " Error : En la Función grabarCallId()", MsgBoxStyle.Critical, "CORDIALPHONE")
        End Try
    End Sub

    ' ##########################################################################################
    '
    '                                  PILA DE PANTALLAS
    '
    ' ##########################################################################################

    ' procedimiento que guarda el indice de la pantalla anterior y ademas refresca el boton volver
    Public Sub guardarPantallaAnterior(ByRef indiceAnterior As Integer)

        pilaPush(indiceAnterior)

        frmVenta.CmdAtras.Enabled = indiceAnterior >= 1

    End Sub

    Public Sub pilaInicializar()
        ' inicializar pila
        pila_cuenta = 0
        ReDim pila(pila_cuenta)
    End Sub

    Public Function pilaPop() As Integer
        If pila_cuenta > 0 Then
            pilaPop = pila(pila_cuenta)
            pila_cuenta = pila_cuenta - 1
            ReDim Preserve pila(pila_cuenta)
        Else
            pilaPop = 0
        End If
    End Function

    Public Function pilaCount() As Integer
        pilaCount = pila_cuenta
    End Function

    Public Sub pilaPush(ByRef elemento As Object)
        pila_cuenta = pila_cuenta + 1
        ReDim Preserve pila(pila_cuenta)
        'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto elemento. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        pila(pila_cuenta) = elemento
    End Sub


    Public Sub buscaValorCombo(ByRef combo As System.Windows.Forms.ComboBox, ByRef buscar As String)
        Dim i As Integer
        Dim indice As Integer

        If Not esVacio(buscar) Then

            ' flag_buscando_en_combo = True

            indice = -1
            With combo
                For i = 0 To .Items.Count - 1
                    .SelectedIndex = i
                    If UCase(.Text) = UCase(buscar) Then
                        indice = i
                        Exit For
                    End If
                Next i
            End With

            ' flag_buscando_en_combo = False

        End If

    End Sub



    Function SoloNumeros(ByVal Keyascii As Short) As Short
        If InStr("1234567890", Chr(Keyascii)) = 0 Then
            SoloNumeros = 0
        Else
            SoloNumeros = Keyascii
        End If
        Select Case Keyascii
            Case 8
                SoloNumeros = Keyascii
            Case 13
                SoloNumeros = Keyascii
        End Select
    End Function

    Public Function ObtieneFonoVenta(ByRef partfono As String, ByRef origenfono As String) As String
        Dim Area As String
        Dim Fono As String
        If Len(origenfono) = 9 Then 'celular
            Area = Left(origenfono, 2)
            Fono = Right(origenfono, 7)
        ElseIf Len(origenfono) = 11 Then  'fono region
            Area = Mid(origenfono, 4, 2)
            Fono = Right(origenfono, 6)
        Else 'santiago
            Area = "2"
            Fono = origenfono
        End If
        If partfono = "area" Then
            ObtieneFonoVenta = Area
        Else
            ObtieneFonoVenta = Fono
        End If
    End Function
    Public Function ObtieneFonoContacto(ByRef partfono As String, ByRef origenfono As String) As String
        Dim AreaContacto As String
        Dim FonoContacto As String
        If Len(origenfono) = 11 And Left(origenfono, 3) = 809 Then 'celular
            AreaContacto = Mid(origenfono, 4, 1)
            FonoContacto = Right(origenfono, 7)
        ElseIf Len(origenfono) = 11 And Left(origenfono, 3) = 123 Then  'fono region
            AreaContacto = Mid(origenfono, 4, 2)
            FonoContacto = Right(origenfono, 6)
        Else 'santiago
            AreaContacto = "2"
            FonoContacto = origenfono
        End If
        If partfono = "area" Then
            ObtieneFonoContacto = AreaContacto
        Else
            ObtieneFonoContacto = FonoContacto
        End If
    End Function
    Public Function ObtenerMes(ByRef mes As Short) As String
        ObtenerMes = ""
        If mes = CDbl("1") Then
            ObtenerMes = "ENERO"
        End If
        If mes = CDbl("2") Then
            ObtenerMes = "FEBRERO"
        End If
        If mes = CDbl("3") Then
            ObtenerMes = "MARZO"
        End If
        If mes = CDbl("4") Then
            ObtenerMes = "ABRIL"
        End If
        If mes = CDbl("5") Then
            ObtenerMes = "MAYO"
        End If
        If mes = CDbl("6") Then
            ObtenerMes = "JUNIO"
        End If
        If mes = CDbl("7") Then
            ObtenerMes = "JULIO"
        End If
        If mes = CDbl("8") Then
            ObtenerMes = "AGOSTO"
        End If
        If mes = CDbl("9") Then
            ObtenerMes = "SEPTIEMBRE"
        End If
        If mes = CDbl("10") Then
            ObtenerMes = "OCTUBRE"
        End If
        If mes = CDbl("11") Then
            ObtenerMes = "NOVIEMBRE"
        End If
        If mes = CDbl("12") Then
            ObtenerMes = "DICIEMBRE"
        End If

    End Function

    Public Function Verificar_Tecla(ByRef Tecla_Presionada As Short) As Short

        If InStr(1, csValNumeros, Chr(Tecla_Presionada)) Or Tecla_Presionada = 8 Then
            Verificar_Tecla = Tecla_Presionada
            'Else
            ' Si no es ninguna de las indicadas retorna 0
            'Verificar_Tecla = 0
        End If

    End Function

    Public Sub Logear_Usuario(ByRef Usuario As String, ByRef estado As Short)
        biGeneral.Loguear_Usuario(Usuario, estado)
    End Sub

    Public Function valFechaNacAdicional(ByRef dia As String, ByRef mes As String, ByRef anio As String) As Boolean

        ' si alguno de los campos es vacio
        Dim fecha_nacimiento As String
        If Val(dia) > 31 Or Val(dia) < 1 Or Val(mes) > 12 Or Val(mes) < 1 Or Val(anio) > Year(Today) Then
            valFechaNacAdicional = False
            Exit Function
        Else
            ' si la concatenacion de los campos da una fecha valida
            fecha_nacimiento = Trim(dia) & "/" & Trim(mes) & "/" & Trim(anio)

            If Not IsDate(fecha_nacimiento) Then
                valFechaNacAdicional = False
                Exit Function
            Else
                ' si esa edad esta dentro del rango permitido
                If edad(CDate(fecha_nacimiento)) < csEdadMinimaAdi Or edad(CDate(fecha_nacimiento)) > csEdadMaxima Then
                    valFechaNacAdicional = False
                    Exit Function
                End If
            End If
        End If

        valFechaNacAdicional = True

    End Function


    Public Sub Tiempo(ByRef respuesta As Short)
        Dim inicio As Date
        Dim ahora As String
        ahora = FechaServidor()
        inicio = Now
        'UPGRADE_WARNING: El comportamiento de DateDiff puede ser diferente. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
        Do While DateDiff(Microsoft.VisualBasic.DateInterval.Second, inicio, Now) - 5
        Loop

    End Sub

    Public Function CaracterValido(ByVal pTecla As String) As String
        'RM / FUNCION PARA VALIDAR INGRESO CARACTER... ACEPTA TODOS EXEPTO *(ASTERISCO) Y '(APOSTROFE)
        Dim lsCad As String
        lsCad = "*'"
        If InStr(lsCad, UCase(Chr(pTecla))) = 0 Then
            CaracterValido = Asc(UCase(Chr(pTecla)))
        End If
    End Function

    Public Sub IngresarDescanso(ByRef Usuario As String, ByRef tipoDescanso As Integer, ByRef descanso As Boolean)
        biGeneral.IngresarDescanso(Usuario, tipoDescanso, descanso)
    End Sub

End Module