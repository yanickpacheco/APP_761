Option Strict Off
Option Explicit On
Friend Class clsSoapNeo

    'Public oSoapClient As New WSSNeotel.WebService

	Public DEVICE As String
	Public IP As String
	Public PORT As String
	Public STATUS As String
	Public SERVER As String
    Public Usuario As String = ""
	Public SAL_CAMPAÑA_DEFAULT As Short
	Public DESCANSO As String
	Public ESTADO_CRM As String
	Public TIEMPO_LOGIN As String
	Public TIEMPO_ADM As String
	Public TIEMPO_LLAMADA As String
	Public Channel As String
	Public CAMPAÑA As Short
	Public CAMPAÑA_ULT As String
	Public COLA As Short
	Public COLA_ULT As String
	Public SUB_ESTADO As String
	Public SUB_ESTADO_ULT As String
	Public DNIS As String
	Public ANI As String
	Public TELEFONO As String
	Public ANI_TELEFONO_ULT As String
	Public TIPO_DISCADOR As String
	Public TIPO_DISCADOR_ULT As String
	Public CRM As Short
	Public BASE As Short
	Public IDCONTACTO As Short
	Public CAMPO_BUSQUEDA As String
	Public IDLLAMADA As String
	Public CLAVE As String
	Public IDLLAMADA_ULT As Short
	Public LLAMANDO As Boolean
	
	Public RutAtento As String
	Public dvAtento As String
    Public AniAtento As String




	Public Function CargarPosicion(ByRef Usuario As String) As String
        On Error GoTo ErrCargarPosicion
		Dim Cadena As String
		Dim strStringNeotel() As String
		Dim i As Short



        Cadena = oSoapClient.Posicion(Usuario)
        '    Cadena = "DEVICE=SIP/CLSH4156|IP=172.16.8.156|PORT=44062|STATUS=OK|SERVER=NEO2|USUARIO=2621|SAL_CAMPAÑA_DEFAULT=109|SUBTIPO_DESCANSO=0|ESTADO_CRM=Available|TIEMPO_LOGIN=14/07/2014 04:01:50 p.m.|TIEMPO_LLAMADA=14/07/2014 05:07:30 p.m.|CHANNEL=SIP/CLSH4156-00005155|CAMPAÑA=109|CAMPAÑA_ULT=109|COLA=3|COLA_ULT=3|SUB_ESTADO=AGENT|SUB_ESTADO_ULT=AGENT|DNIS=9992|ANI=50038|TELEFONO=|ANI_TELEFONO_ULT=50038|TIPO_LLAMADA=NotSet|TIPO_LLAMADA_ULT=NotSet|ORIGEN_LLAMADA=|ORIGEN_LLAMADA_ULT=|DIRECCION=ENTRANTE|DIRECCION_ULT=ENTRANTE|CRM=0|BASE=0|IDCONTACTO=0|DATA=|CLAVE=RUT=6007453;ANI=223126917;DV=4;|CAMPO_BUSQUEDA=|IDLLAMADA=26396268|IDLLAMADA_ULT=26396268|CONFERENCIA=|GRABANDO=SI|GRABACION=50038-26396268-201407141707.mp3|TELEFONO_DESVIO=|SAL_TIPO_DISCADOR=Manual|SAL_CRM=0|SAL_BASE=0|CANALES_ASOCIADOS=¦¦Zap/126-1|"
        '    Cadena = "DEVICE=SIP/rmartnez_cs|IP=172.16.2.111|PORT=33182|STATUS=OK|SERVER=NEO1|USUARIO=8000|INICIO_LOGIN=14/07/2014 01:56:43 p.m.|SAL_CAMPAÑA_DEFAULT=0|INICIO_LOGIN_CAMPAÑA=01/01/2000 12:00:00 a.m.|CHANNEL=|SUB_ESTADO=|SUB_ESTADO_ULT=|RINGING=01/01/0001 12:00:00 a.m.|DIALING=01/01/0001 12:00:00 a.m.|AGENT=01/01/0001 12:00:00 a.m.|TIEMPO_LLAMADA=14/07/2014 01:59:38 p.m.|SUBTIPO_DESCANSO=0|INICIO_DESCANSO=01/01/2000 12:00:00 a.m.|TIEMPO_DESCANSO=01/01/2000 12:00:00 a.m.|ESTADO_CRM=Available|INICIO_CRM=01/01/2000 12:00:00 a.m.|CAMPAÑA=0|CAMPAÑA_ULT=0|COLA=0|COLA_ULT=0|DNIS=|ANI=|TELEFONO=|ANI_TELEFONO_ULT=|TIPO_LLAMADA=NotSet|TIPO_LLAMADA_ULT=NotSet|ORIGEN_LLAMADA=|ORIGEN_LLAMADA_ULT=|DIRECCION=|DIRECCION_ULT=|CRM=0|BASE=0|IDCONTACTO=0|DATA=|CLAVE=|CAMPO_BUSQUEDA=|IDAGENDA=|IDLLAMADA=0|IDLLAMADA_ULT=0|CONFERENCIA=|GRABANDO=NO|GRABACION=|TELEFONO_DESVIO=|SAL_TIPO_DISCADOR=NotSet|SAL_CRM=0|SAL_BASE=0|CANALES_ASOCIADOS=|"		
        '              DEVICE=SIP/soporte_mok|IP=0.0.0.0|PORT=0|STATUS=|SERVER=|USUARIO=|INICIO_LOGIN=01/01/2000 12:00:00 a.m.|SAL_CAMPAÑA_DEFAULT=0|INICIO_LOGIN_CAMPAÑA=01/01/2000 12:00:00 a.m.|CHANNEL=|SUB_ESTADO=|SUB_ESTADO_ULT=|RINGING=01/01/0001 12:00:00 a.m.|DIALING=01/01/0001 12:00:00 a.m.|AGENT=01/01/0001 12:00:00 a.m.|TIEMPO_LLAMADA=01/01/2000 12:00:00 a.m.|SUBTIPO_DESCANSO=0|INICIO_DESCANSO=01/01/2000 12:00:00 a.m.|TIEMPO_DESCANSO=01/01/2000 12:00:00 a.m.|ESTADO_CRM=Available|INICIO_CRM=01/01/2000 12:00:00 a.m.|CAMPAÑA=0|CAMPAÑA_ULT=0|COLA=0|COLA_ULT=0|DNIS=|ANI=|TELEFONO=|ANI_TELEFONO_ULT=|TIPO_LLAMADA=NotSet|TIPO_LLAMADA_ULT=NotSet|ORIGEN_LLAMADA=|ORIGEN_LLAMADA_ULT=|DIRECCION=|DIRECCION_ULT=|CRM=0|BASE=0|IDCONTACTO=0|DATA=|CLAVE=|CAMPO_BUSQUEDA=|IDAGENDA=|IDLLAMADA=0|IDLLAMADA_ULT=0|CONFERENCIA=|GRABANDO=NO|GRABACION=|TELEFONO_DESVIO=|SAL_TIPO_DISCADOR=NotSet|SAL_CRM=0|SAL_BASE=0|CANALES_ASOCIADOS=|
        If Cadena <> "" Then
            'Dim vpPosicion As New SoapNeo
            strStringNeotel = Split(Cadena, "|")


            For i = LBound(strStringNeotel) To UBound(strStringNeotel) - 1

                If UCase(Split(strStringNeotel(i), "=")(0)) = "IP" Then
                    vpPosicion.IP = Split(strStringNeotel(i), "=")(1)
                End If

                If UCase(Split(strStringNeotel(i), "=")(0)) = "STATUS" Then
                    vpPosicion.STATUS = Split(strStringNeotel(i), "=")(1)
                End If

                If UCase(Split(strStringNeotel(i), "=")(0)) = "SERVER" Then
                    vpPosicion.SERVER = Split(strStringNeotel(i), "=")(1)
                End If

                If UCase(Split(strStringNeotel(i), "=")(0)) = "USUARIO" Then
                    vpPosicion.Usuario = Split(strStringNeotel(i), "=")(1)
                End If

                If UCase(Split(strStringNeotel(i), "=")(0)) = "IDLLAMADA" Then
                    vpPosicion.IDLLAMADA = IIf(CDbl(Split(strStringNeotel(i), "=")(1)) = 0, "", Right("0000000000" & Split(strStringNeotel(i), "=")(1), 10))
                End If

                If UCase(Split(strStringNeotel(i), "=")(0)) = "CLAVE" Then
                    vpPosicion.CLAVE = Replace(strStringNeotel(i), "CLAVE=", "")
                End If

                If UCase(Split(strStringNeotel(i), "=")(0)) = "TELEFONO" Then
                    vpPosicion.TELEFONO = Split(strStringNeotel(i), "=")(1)
                End If

                If UCase(Split(strStringNeotel(i), "=")(0)) = "ANI" Then
                    vpPosicion.ANI = Split(strStringNeotel(i), "=")(1)
                End If

                If UCase(Split(strStringNeotel(i), "=")(0)) = "DNIS" Then
                    vpPosicion.DNIS = Split(strStringNeotel(i), "=")(1)
                End If

                If UCase(Split(strStringNeotel(i), "=")(0)) = "COLA" Then
                    vpPosicion.COLA = CShort(Split(strStringNeotel(i), "=")(1))
                End If

                'If UCase(Split(strStringNeotel(i), "=")(0)) = "CAMPAÑA" Then
                '    vpPosicion.CAMPAÑA = CShort(Split(strStringNeotel(i), "=")(1))
                'End If

                If UCase(Split(strStringNeotel(i), "=")(0)) = "IDCONTACTO" Then
                    vpPosicion.IDCONTACTO = CShort(Split(strStringNeotel(i), "=")(1))
                End If

                If UCase(Split(strStringNeotel(i), "=")(0)) = "SAL_CAMPAÑA_DEFAULT" Then
                    vpPosicion.SAL_CAMPAÑA_DEFAULT = CShort(Split(strStringNeotel(i), "=")(1))
                End If

                If UCase(Split(strStringNeotel(i), "=")(0)) = "DESCANSO" Then
                    vpPosicion.DESCANSO = Split(strStringNeotel(i), "=")(1)
                End If

                If UCase(Split(strStringNeotel(i), "=")(0)) = "TIEMPO_ADM" Then
                    vpPosicion.TIEMPO_ADM = Split(strStringNeotel(i), "=")(1)
                End If

                If UCase(Split(strStringNeotel(i), "=")(0)) = "TIEMPO_LOGIN" Then
                    vpPosicion.TIEMPO_LOGIN = Split(strStringNeotel(i), "=")(1)
                End If

                If UCase(Split(strStringNeotel(i), "=")(0)) = "TIEMPO_LLAMADA" Then
                    vpPosicion.TIEMPO_LLAMADA = Split(strStringNeotel(i), "=")(1)
                End If


            Next
            CargarPosicion = vpPosicion.Usuario
        Else
            MsgBox("Debe loguearse en XLite antes de ingresar al CRM", MsgBoxStyle.Critical, "Atención")
            CargarPosicion = ""
        End If
        oSoapClient.CRM_Disponible(Usuario)

ErrCargarPosicion:
        If Err.Description <> "" Then
            Select Case Err.Number
                Case CInt("-2147467259")
                    MsgBox("Debe loguearse en XLite antes de ingresar al CRM", MsgBoxStyle.Critical, "Atención")
                Case Else
                    MsgBox(Err.Description & vbNewLine & " Se produjo un problema al Loguear al Usuario Intentelo Nuevamente", MsgBoxStyle.Critical, "CORDIALPHONE")
            End Select
            Err.Clear()
        End If
    End Function
	Public Sub LogoutTelefonia(ByRef Usuario As Object)		
        oSoapClient.Logout(Usuario)
    End Sub
	Public Function Login_Campania(ByRef Usuario As String, ByRef Campania As Short) As Boolean
		On Error GoTo ErrLogin_Campania		
		oSoapClient.Login_Campaña(Usuario, Campania)
		Login_Campania = True
ErrLogin_Campania: 
		If Err.Description <> "" Then
			Login_Campania = False
			MsgBox(Err.Description, MsgBoxStyle.Critical, "Atención")
			Err.Clear()
		End If
	End Function
    Public Sub Logout_Campania(ByRef Usuario As String, ByRef Campania As Short)
        oSoapClient.Logout_Campaña(Usuario)
    End Sub
	Public Sub Discar(ByRef Usuario As String, ByRef NumTelefono As String)

		Dim Fono As String
        Fono = NumTelefono
        'If Len(NumTelefono) = 9 Then
        '    Fono = NumTelefono
        'ElseIf Len(NumTelefono) > 9 And Left(NumTelefono, 2) = "09" Then
        '    Fono = "09" & Mid(NumTelefono, 4, Len(NumTelefono))
        'ElseIf Len(NumTelefono) > 9 And Left(NumTelefono, 3) = "123" Then
        '    Fono = "0" & Mid(NumTelefono, 4, Len(NumTelefono))
        'Else
        '    MsgBox("No es posible Generar el Llamado Comuniquece con Supervisor ", MsgBoxStyle.Information, "CORDIALPHONE")
        '    Exit Sub
        'End If


        oSoapClient.Discar(Usuario, Fono)
        vpPosicion.CargarPosicion((vpPosicion.Usuario))
        Dim Contador As Integer = 0
        Do While vpPosicion.IDLLAMADA = ""
            System.Threading.Thread.Sleep(1000)
            vpPosicion.CargarPosicion((vpPosicion.Usuario))
            Contador += 1
            If Contador = 5 Then
                Exit Do
            End If
        Loop

		vpPosicion.LLAMANDO = True
	End Sub
	Public Sub Cortar(ByRef Usuario As String)		
        oSoapClient.Cortar(Usuario)
        vpPosicion.LLAMANDO = False
    End Sub
    Public Function EvaluaEstado(ByRef Usuario As String) As Boolean
        Dim IDLLAMADA As String
        Dim Cadena As String
        EvaluaEstado = True
        Cadena = oSoapClient.Posicion(Usuario)

        If Cadena <> "" Then
            IDLLAMADA = IIf(CDbl(Mid(Cadena, InStr(1, Cadena, "IDLLAMADA=") + 10, InStr(1, Mid(Cadena, InStr(1, Cadena, "IDLLAMADA=") + 10, 15), "|") - 1)) = 0, "", Right("0000000000" & Mid(Cadena, InStr(1, Cadena, "IDLLAMADA=") + 10, InStr(1, Mid(Cadena, InStr(1, Cadena, "IDLLAMADA=") + 10, 15), "|") - 1), 10))
            If IDLLAMADA = "" And vpPosicion.LLAMANDO = True Then
                EvaluaEstado = False
            End If
        End If
    End Function

    Public Sub ActivarDescanso(ByVal Usuario As String, ByVal descanso As Boolean)
        oSoapClient.Descanso(Usuario, descanso)
    End Sub
End Class