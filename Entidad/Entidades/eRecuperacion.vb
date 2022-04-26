Public Class eRecuperacion
    Private _R_ID As Integer
    Private _C_ID As Integer
    Private _R_Estado As String
    Private _R_Fecha As String
    Private _R_Hora As String
    Private _R_Fecha_Vta As String
    Private _R_Hora_Vta As String
    Private _R_IdLlamada As String
    Private _R_Intento As Integer
    Private _R_Ejecutivo As String
    Private _R_Fecha_Agen As String
    Private _R_Hora_Agen As String
    Private _R_Obs_Agen As String
    Private _R_Contacto As String
    Private _R_No_Contacto As String
    Private _R_Contacto_Con As String
    Private _R_Contacto_Flujo As String
    Private _R_Motivo_No_Interesa As String
    Private _R_Reconoce As String
    Private _R_Venta As Integer

    Public Sub New()
        _R_ID = 0
        _C_ID = 0
        _R_Estado = ""
        _R_Fecha = ""
        _R_Hora = ""
        _R_Fecha_Vta = ""
        _R_Hora_Vta = ""
        _R_IdLlamada = ""
        _R_Intento = 0
        _R_Ejecutivo = ""
        _R_Fecha_Agen = ""
        _R_Hora_Agen = ""
        _R_Obs_Agen = ""
        _R_Contacto = ""
        _R_No_Contacto = ""
        _R_Contacto_Con = ""
        _R_Contacto_Flujo = ""
        _R_Motivo_No_Interesa = ""
        _R_Reconoce = ""
        _R_Venta = 0

    End Sub

    Public Property R_ID As Integer
        Get
            Return _R_ID
        End Get
        Set(value As Integer)
            _R_ID = value
        End Set
    End Property

    Public Property C_ID As Integer
        Get
            Return _C_ID
        End Get
        Set(value As Integer)
            _C_ID = value
        End Set
    End Property

    Public Property R_Estado As String
        Get
            Return _R_Estado
        End Get
        Set(value As String)
            _R_Estado = value
        End Set
    End Property

    Public Property R_Fecha As String
        Get
            Return _R_Fecha
        End Get
        Set(value As String)
            _R_Fecha = value
        End Set
    End Property

    Public Property R_Hora As String
        Get
            Return _R_Hora
        End Get
        Set(value As String)
            _R_Hora = value
        End Set
    End Property

    Public Property R_Fecha_Vta As String
        Get
            Return _R_Fecha_Vta
        End Get
        Set(value As String)
            _R_Fecha_Vta = value
        End Set
    End Property

    Public Property R_Hora_Vta As String
        Get
            Return _R_Hora_Vta
        End Get
        Set(value As String)
            _R_Hora_Vta = value
        End Set
    End Property

    Public Property R_IdLlamada As String
        Get
            Return _R_IdLlamada
        End Get
        Set(value As String)
            _R_IdLlamada = value
        End Set
    End Property

    Public Property R_Intento As Integer
        Get
            Return _R_Intento
        End Get
        Set(value As Integer)
            _R_Intento = value
        End Set
    End Property

    Public Property R_Ejecutivo As String
        Get
            Return _R_Ejecutivo
        End Get
        Set(value As String)
            _R_Ejecutivo = value
        End Set
    End Property

    Public Property R_Fecha_Agen As String
        Get
            Return _R_Fecha_Agen
        End Get
        Set(value As String)
            _R_Fecha_Agen = value
        End Set
    End Property

    Public Property R_Hora_Agen As String
        Get
            Return _R_Hora_Agen
        End Get
        Set(value As String)
            _R_Hora_Agen = value
        End Set
    End Property

    Public Property R_Obs_Agen As String
        Get
            Return _R_Obs_Agen
        End Get
        Set(value As String)
            _R_Obs_Agen = value
        End Set
    End Property

    Public Property R_Contacto As String
        Get
            Return _R_Contacto
        End Get
        Set(value As String)
            _R_Contacto = value
        End Set
    End Property

    Public Property R_No_Contacto As String
        Get
            Return _R_No_Contacto
        End Get
        Set(value As String)
            _R_No_Contacto = value
        End Set
    End Property

    Public Property R_Contacto_Con As String
        Get
            Return _R_Contacto_Con
        End Get
        Set(value As String)
            _R_Contacto_Con = value
        End Set
    End Property

    Public Property R_Contacto_Flujo As String
        Get
            Return _R_Contacto_Flujo
        End Get
        Set(value As String)
            _R_Contacto_Flujo = value
        End Set
    End Property

    Public Property R_Motivo_No_Interesa As String
        Get
            Return _R_Motivo_No_Interesa
        End Get
        Set(value As String)
            _R_Motivo_No_Interesa = value
        End Set
    End Property

    Public Property R_Reconoce As String
        Get
            Return _R_Reconoce
        End Get
        Set(value As String)
            _R_Reconoce = value
        End Set
    End Property

    Public Property R_Venta As Integer
        Get
            Return _R_Venta
        End Get
        Set(value As Integer)
            _R_Venta = value
        End Set
    End Property
End Class
