Public Class eLlamada

    Private _LL_ID As Integer
    Private _C_ID As String
    Private _LL_CallId As String
    Private _LL_Fecha As String
    Private _LL_Hora As String
    Private _LL_Ejecutivo As String
    Private _LL_Hora_Inicio As String
    Private _LL_Hora_Termino As String
    Private _LL_Fono_Contacto As String

    Public Sub New(lL_ID As Integer, C_ID As String, LL_CallId As String, LL_Fecha As String, LL_Hora As String, LL_Ejecutivo As String, LL_Hora_Inicio As String, LL_Hora_Termino As String, LL_Fono_Contacto As String)

        Me.LL_ID = lL_ID
        Me.C_ID = C_ID
        Me.LL_CallId = LL_CallId
        Me.LL_Fecha = LL_Fecha
        Me.LL_Hora = LL_Hora
        Me.LL_Ejecutivo = LL_Ejecutivo
        Me.LL_Hora_Inicio = LL_Hora_Inicio
        Me.LL_Hora_Termino = LL_Hora_Termino
        Me._LL_Fono_Contacto = LL_Fono_Contacto
    End Sub

    Public Sub New()
        LL_ID = 0
        C_ID = 0
        LL_CallId = ""
        LL_Fecha = ""
        LL_Hora = ""
        LL_Ejecutivo = ""
        LL_Hora_Inicio = ""
        LL_Hora_Termino = ""
        _LL_Fono_Contacto = ""

    End Sub


    Public Property LL_ID As Integer
        Get
            Return _LL_ID
        End Get
        Set(value As Integer)
            _LL_ID = value
        End Set
    End Property

    Public Property C_ID As String
        Get
            Return _C_ID
        End Get
        Set(value As String)
            _C_ID = value
        End Set
    End Property

    Public Property LL_CallId As String
        Get
            Return _LL_CallId
        End Get
        Set(value As String)
            _LL_CallId = value
        End Set
    End Property

    Public Property LL_Fecha As String
        Get
            Return _LL_Fecha
        End Get
        Set(value As String)
            _LL_Fecha = value
        End Set
    End Property

    Public Property LL_Hora As String
        Get
            Return _LL_Hora
        End Get
        Set(value As String)
            _LL_Hora = value
        End Set
    End Property

    Public Property LL_Ejecutivo As String
        Get
            Return _LL_Ejecutivo
        End Get
        Set(value As String)
            _LL_Ejecutivo = value
        End Set
    End Property

    Public Property LL_Hora_Inicio As String
        Get
            Return _LL_Hora_Inicio
        End Get
        Set(value As String)
            _LL_Hora_Inicio = value
        End Set
    End Property

    Public Property LL_Hora_Termino As String
        Get
            Return _LL_Hora_Termino
        End Get
        Set(value As String)
            _LL_Hora_Termino = value
        End Set
    End Property

    Public Property LL_Fono_Contacto As String
        Get
            Return _LL_Fono_Contacto
        End Get
        Set(value As String)
            _LL_Fono_Contacto = value
        End Set
    End Property
End Class
