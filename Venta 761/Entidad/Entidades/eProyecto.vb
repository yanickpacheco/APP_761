Imports Entidad

Public Class eProyecto

    Private _idCRM As Int32
    Private _nombre As String
    Private _codigo As String
    Private _idtipoCampania As Int32
    Private _adicionales As Boolean
    Private _beneficiarios As Boolean
    Private _medioPago As Boolean
    Private _servidorBDD As String
    Private _idTelefonia As Integer

    Public Sub New(idCRM As Integer, nombre As String, codigo As String, idtipoCampania As Integer, adicionales As Boolean, beneficiarios As Boolean, medioPago As Boolean, servidorBDD As String, idTelefonia As Integer)
        _idCRM = idCRM
        _nombre = nombre
        _codigo = codigo
        _idtipoCampania = idtipoCampania
        _adicionales = adicionales
        _beneficiarios = beneficiarios
        _medioPago = medioPago
        Me.ServidorBDD = servidorBDD
        Me.IdTelefonia = idTelefonia
    End Sub

    Public Sub New()
        _idCRM = 0
        _nombre = ""
        _codigo = ""
        _idtipoCampania = 0
        _adicionales = False
        _beneficiarios = False
        _medioPago = False
        Me.ServidorBDD = ""
        Me.IdTelefonia = 0
    End Sub


    Public Property idCRM() As Int32
        Get
            Return _idCRM
        End Get
        Set(ByVal value As Int32)
            _idCRM = value
        End Set
    End Property

    Public Property nombre() As String
        Get
            Return _nombre
        End Get
        Set(ByVal value As String)
            _nombre = value
        End Set
    End Property

    Public Property codigo() As String

        Get
            Return _codigo
        End Get
        Set(ByVal value As String)
            _codigo = value
        End Set
    End Property



    Public Property idtipoCampania() As Int32
        Get
            Return _idtipoCampania
        End Get
        Set(ByVal value As Int32)
            _idtipoCampania = value
        End Set
    End Property

    Public Property adicionales() As Boolean
        Get
            Return _adicionales
        End Get
        Set(ByVal value As Boolean)
            _adicionales = value
        End Set
    End Property

    Public Property beneficiarios() As Boolean
        Get
            Return _beneficiarios
        End Get
        Set(ByVal value As Boolean)
            _beneficiarios = value
        End Set
    End Property

    Public Property medioPago() As Boolean
        Get
            Return _medioPago
        End Get
        Set(ByVal value As Boolean)
            _medioPago = value
        End Set
    End Property

    Public Property ServidorBDD As String
        Get
            Return _servidorBDD
        End Get
        Set(value As String)
            _servidorBDD = value
        End Set
    End Property

    Public Property IdTelefonia As Integer
        Get
            Return _idTelefonia
        End Get
        Set(value As Integer)
            _idTelefonia = value
        End Set
    End Property
End Class
