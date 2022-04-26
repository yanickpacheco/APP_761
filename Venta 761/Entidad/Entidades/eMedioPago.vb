Public Class eMedioPago

    Private _idMedioPago As Integer
    Private _C_ID As Integer
    Private _MedioPago As String
    Private _Banco As String
    Private _Tarjeta As String
    Private _NumeroMedioPago As String
    Private _MesVencimiento As String
    Private _AñoVencimiento As String

    Private _otroMedioPago As Boolean
    Public Property otroMedioPago() As Boolean
        Get
            Return _otroMedioPago
        End Get
        Set(ByVal value As Boolean)
            _otroMedioPago = value
        End Set
    End Property

    'Public Sub New(idMedioPago As Integer, c_ID As Integer, medioPago As String, banco As String, tarjeta As String, numeroMedioPago As String, mesVencimiento As String, añoVencimiento As String)
    '    idMedioPago = idMedioPago
    '    Me.C_ID = c_ID
    '    Me.MedioPago = medioPago
    '    Me.Banco = banco
    '    Me.Tarjeta = tarjeta
    '    Me.NumeroMedioPago = numeroMedioPago
    '    Me.MesVencimiento = mesVencimiento
    '    Me.AñoVencimiento = añoVencimiento
    'End Sub

    Public Property idMedioPago As Integer
        Get
            Return _idMedioPago
        End Get
        Set(value As Integer)
            _idMedioPago = value
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

    Public Property MedioPago As String
        Get
            Return _MedioPago
        End Get
        Set(value As String)
            _MedioPago = value
        End Set
    End Property

    Public Property Banco As String
        Get
            Return _Banco
        End Get
        Set(value As String)
            _Banco = value
        End Set
    End Property

    Public Property Tarjeta As String
        Get
            Return _Tarjeta
        End Get
        Set(value As String)
            _Tarjeta = value
        End Set
    End Property

    Public Property NumeroMedioPago As String
        Get
            Return _NumeroMedioPago
        End Get
        Set(value As String)
            _NumeroMedioPago = value
        End Set
    End Property

    Public Property MesVencimiento As String
        Get
            Return _MesVencimiento
        End Get
        Set(value As String)
            _MesVencimiento = value
        End Set
    End Property

    Public Property AñoVencimiento As String
        Get
            Return _AñoVencimiento
        End Get
        Set(value As String)
            _AñoVencimiento = value
        End Set
    End Property
End Class
