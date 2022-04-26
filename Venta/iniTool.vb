Option Strict Off
Option Explicit On
Friend Class iniTool
	Private mIniFileName As String
	
	'UPGRADE_ISSUE: No se admite la declaración de un parámetro 'As Any'. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
	'UPGRADE_ISSUE: No se admite la declaración de un parámetro 'As Any'. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	
	
	
	Public Property FileName() As String
		Get
			'On Error Resume Next
			FileName = mIniFileName
		End Get
		Set(ByVal Value As String)
			'On Error Resume Next
			Dim NF As Integer
			'UPGRADE_WARNING: Dir tiene un nuevo comportamiento. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If Dir(Value, FileAttribute.Normal) = "" Then
				NF = FreeFile
				FileOpen(NF, Value, OpenMode.Output)
				FileClose(NF)
			End If
			mIniFileName = Value
		End Set
	End Property
	
	
	Public Function GetValue(ByVal Section As String, ByVal Key As String, Optional ByVal DefaultValue As String = "") As String
		On Error GoTo xERR
		Dim Value, retval As String
		Dim X As Short
		retval = New String(Chr(0), 255)
		X = GetPrivateProfileString(Section, Key, DefaultValue, retval, Len(retval), mIniFileName)
		GetValue = Trim(Left(retval, X))
		Exit Function
xERR: 
		GetValue = DefaultValue
	End Function
	
	Public Function WriteValue(ByVal Section As String, ByVal Key As String, ByVal Value As String) As Boolean
		On Error GoTo xERR
		Dim X As Short
		X = WritePrivateProfileString(Section, Key, Value, mIniFileName)
		If X <> 0 Then WriteValue = True
		Exit Function
xERR: 
	End Function
	
	Public Function GetAllSections() As Collection
		Dim Value, retval As String
		Dim X As Short
		Dim s() As String
		Dim i As Short
		retval = New String(Chr(0), 255)
		X = GetPrivateProfileString(vbNullString, "", "", retval, Len(retval), mIniFileName)
		Value = Trim(Left(retval, X))
		s = Split(Value, Chr(0))
		GetAllSections = New Collection
		With GetAllSections
			For i = LBound(s) To UBound(s)
				If s(i) <> "" Then .Add(s(i))
			Next 
		End With
	End Function
	
	Public Function GetAllKeys(ByVal Section As String) As Collection
		Dim Value, retval As String
		Dim X As Short
		Dim s() As String
		Dim i As Short
		retval = New String(Chr(0), 255)
		X = GetPrivateProfileString(Section, vbNullString, "", retval, Len(retval), mIniFileName)
		Value = Trim(Left(retval, X))
		s = Split(Value, Chr(0))
		GetAllKeys = New Collection
		With GetAllKeys
			For i = LBound(s) To UBound(s)
				If s(i) <> "" Then .Add(s(i))
			Next 
		End With
	End Function
	
	Public Function DeleteSection(ByVal Section As String) As Boolean
		On Error GoTo xERR
		Dim X As Short
		X = WritePrivateProfileString(Section, vbNullString, "", mIniFileName)
		If X <> 0 Then DeleteSection = True
		Exit Function
xERR: 
	End Function
	
	Public Function DeleteKey(ByVal Section As String, ByVal Key As String) As Boolean
		On Error GoTo xERR
		Dim X As Short
		X = WritePrivateProfileString(Section, Key, vbNullString, mIniFileName)
		If X <> 0 Then DeleteKey = True
		Exit Function
xERR: 
	End Function
End Class