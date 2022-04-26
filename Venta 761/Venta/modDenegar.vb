Option Strict Off
Option Explicit On
Module modDenegar
	
	Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Integer, ByVal wFlag As Integer) As Integer
	'UPGRADE_ISSUE: No se admite la declaración de un parámetro 'As Any'. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As String) As Integer
	Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Integer) As Integer
	Public Declare Function GetDesktopWindow Lib "user32" () As Integer
	
	
	Public Declare Function GetWindowTextLength Lib "user32"  Alias "GetWindowTextLengthA"(ByVal hWnd As Integer) As Integer
	Public Declare Function GetWindowText Lib "user32"  Alias "GetWindowTextA"(ByVal hWnd As Integer, ByVal lpString As String, ByVal cch As Integer) As Integer
	
	' GetWindow() Constants
	Public Const GW_HWNDFIRST As Short = 0
	Public Const GW_HWNDNEXT As Short = 2
	Public Const GW_CHILD As Short = 5
	
	Public Const SC_MINIMIZE As Integer = &HF020
	Public Const SC_CLOSE As Integer = &HF060
	Public Const WM_SYSCOMMAND As Integer = &H112
	
	
	Public Function WindowTitle(ByVal hWnd As Integer) As String
		'Devuelve el título de una ventana, según el hWnd indicado
		'
		Dim sTitulo As String
		Dim lenTitulo As Integer
		Dim ret As Integer
        WindowTitle = ""
		'Leer la longitud del título de la ventana
		lenTitulo = GetWindowTextLength(hWnd)
		If lenTitulo > 0 Then
			lenTitulo = lenTitulo + 1
			sTitulo = New String(Chr(0), lenTitulo)
			'Leer el título de la ventana
			ret = GetWindowText(hWnd, sTitulo, lenTitulo)
			WindowTitle = Left(sTitulo, ret)
		End If
	End Function
	
	'---------------------------------------------------------------------------------------
	' Procedure : denegarAplicacion
	' DateTime  : 18/07/2006 16:49
	' Author    : emunoz
	' Purpose   : busca en la lista de las ventanas activas si la aplicacion que contenga en
	'           la frase aBuscar en su titulo este, y si esta la MINIMIZA/CIERRA
	'---------------------------------------------------------------------------------------
	'
	Public Sub denegarAplicacion(ByRef aBuscar As String)
		
		Dim sTitulo As String
		Dim hWnd As Integer
		
		' primera ventana (toda la pantalla)
		hWnd = GetWindow(GetDesktopWindow(), GW_CHILD)
		
		'Recorrer el resto de las ventanas
		Do While hWnd <> 0
			'Si la ventana es visible
			If IsWindowVisible(hWnd) Then
				'Leer el caption de la ventana
				sTitulo = WindowTitle(hWnd)
				If Len(sTitulo) Then
					' vemos si esta ventana tiene el titulo aBuscar
					If InStr(1, sTitulo, aBuscar, CompareMethod.Text) Then
						' encontrada! entonces lo cerramos
						'Call SendMessage(hWnd, WM_SYSCOMMAND, SC_CLOSE, 0)
						Call SendMessage(hWnd, WM_SYSCOMMAND, SC_MINIMIZE, 0)
						Exit Do ' quebrar del while
					End If
				End If
			End If
			'Siguiente ventana
			hWnd = GetWindow(hWnd, GW_HWNDNEXT)
		Loop 
	End Sub
End Module