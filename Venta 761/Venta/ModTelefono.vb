Option Strict Off
Option Explicit On
Module Mod_Telefono
	
	'Estudiar funciones de coneccion
	
	Public Structure tCICObject
		Dim p_Type As Integer
		Dim p_Id As Integer
	End Structure
	
	'Public icConexion As EICServer2
	'Public icQueue As EICQueue
	'Public icLlamadoActual As EICCallObject
	'Public icLlamadoActual2 As EICCallObject
	
	Public objetosCola() As tCICObject
	
	Public ic_llamando As Boolean
	Public ic_callid As String
	Public ic_ultimo_fono_llamado As String
	'Public Function conectarCIC(servidor As String, Usuario As String, password As String, estacion As String) As Boolean
	'
	'    Dim resultado As Boolean
	'
	'    resultado = True
	'
	'    On Error GoTo conectarCIC_salida
	'
	'    'realizamos coneccion con el servidor Cic
	'    Set icConexion = New EICServer2
	'
	'    icConexion.Connect "SRV-CIC23", servidor, Usuario, password, estacion, False
	'
	'    'icConexion.Connect "123", servidor, usuario, password, estacion, False
	'    'inicializar cola de llamados
	'
	'    Set icQueue = New EICQueue
	'    icQueue.Connect QT_USER, icConexion.User
	'
	'    GoTo conectarCIC_salida
	'
	'conectarCIC_salida:
	'
	'    If Err.Number <> 0 Then
	'        MsgBox "Ha ocurrido un error al conectar al servidor de telefonía." & vbCrLf & vbCrLf & Err.Description, vbCritical, "conectarCIC"
	'        resultado = False
	'    End If
	'
	'    conectarCIC = resultado
	'
	'End Function
	
	
	Public Function llamar(ByRef Fono As String) As Boolean
		On Error GoTo llamar_salida

        vpPosicion.Discar((vpPosicion.Usuario), Fono)
        WS_CALL_ID = vpPosicion.IDLLAMADA
        llamar = True

llamar_salida:
        If Err.Number <> 0 Then
            ic_llamando = False
            MsgBox("Ocurrio un error al intentar realizar el llamado." & vbCrLf & Err.Description, MsgBoxStyle.Critical, csNombreAplicacion)
        End If
		
	End Function
	Public Function colgar(ByRef ConGestion As Boolean) As Boolean
		On Error GoTo colgar_salida
		If db_central = 1 Or db_central = 3 Then
			'        icLlamadoActual.Disconnect ""
			'        Set icLlamadoActual = Nothing
		ElseIf db_central = 4 Then 
			vpPosicion.Cortar((vpPosicion.Usuario))
		End If
		GoTo colgar_salida
colgar_salida: 
		ic_llamando = False
		colgar = True
		
	End Function
	Public Sub Mute()
		If db_central = 1 Or db_central = 3 Then
			If ic_llamando Then
				'icLlamadoActual.Mute ""
			End If
		Else
			MsgBox("Opción no habilitada", MsgBoxStyle.Information, "Aviso")
		End If
	End Sub
	
	Public Function EstadoLLamada() As String
        EstadoLLamada = ""
        If db_central = 1 Or db_central = 3 Then
            'EstadoLLamada = icLlamadoActual.State & " " & icLlamadoActual.StateString            
        End If
	End Function
	Public Sub Hold()
		If db_central = 1 Or db_central = 3 Then
			If ic_llamando Then
				'icLlamadoActual.Hold "", ""
			End If
		Else
			MsgBox("Opción no habilitada", MsgBoxStyle.Information, "Aviso")
		End If
	End Sub
	Public Sub Pickup()
		If db_central = 1 Or db_central = 3 Then
			'        If icLlamadoActual.StateString = "Held" Then
			'            icLlamadoActual.Pickup ""
			'        End If
		Else
			MsgBox("Opción no habilitada", MsgBoxStyle.Information, "Aviso")
		End If
	End Sub
	
	Public Sub llamadaInstalacion(ByRef Fono As String)
		'Set icLlamadoActual2 = New EICCallObject
		
		'icLlamadoActual2.Dial Fono, 1
		
		'icLlamadoActual.ExtendedDial fono, 1, False, True
		
	End Sub
	
	Public Sub Transferencia()
		If ic_llamando Then
			'icLlamadoActual.ConsultTransfer icLlamadoActual2.Id, ""
		End If
	End Sub
	
	Public Function colgarAreaValidacion() As Boolean
		
		Dim resultado As Boolean
		
		On Error GoTo colgar_salida
		
		'icLlamadoActual2.Disconnect ""
		
		'Set icLlamadoActual2 = Nothing
		
		GoTo colgar_salida
		
colgar_salida: 
		ic_llamando = False
		colgarAreaValidacion = True
		
	End Function
End Module