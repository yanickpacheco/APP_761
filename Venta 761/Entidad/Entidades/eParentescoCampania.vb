Public Class eTipoParentesco
    Private _idParentescoCampania As Int16
    Private _idParentesco As Int16
    Private _nombreParentesco As String
    Private _calCodigo As Int16
    Private _edadMin As Int16
    Private _edadMax As Int16
    Private _idTipoPersonaAgregado As Int16
    Private _idPlan As Int64
    Public Property idPlan() As Int64
        Get
            Return _idPlan
        End Get
        Set(ByVal value As Int64)
            _idPlan = value
        End Set
    End Property

    Public Property nombreParentesco() As String
        Get
            Return _nombreParentesco
        End Get
        Set(ByVal value As String)
            _nombreParentesco = value
        End Set
    End Property
    Public Property idTipoPersonaAgregado() As Int16
        Get
            Return _idTipoPersonaAgregado
        End Get
        Set(ByVal value As Int16)
            _idTipoPersonaAgregado = value
        End Set
    End Property

    Public Property edadMax() As Int16
        Get
            Return _edadMax
        End Get
        Set(ByVal value As Int16)
            _edadMax = value
        End Set
    End Property

    Public Property edadMin() As Int16
        Get
            Return _edadMin
        End Get
        Set(ByVal value As Int16)
            _edadMin = value
        End Set
    End Property

    Public Property calCodigo() As Int16
        Get
            Return _calCodigo
        End Get
        Set(ByVal value As Int16)
            _calCodigo = value
        End Set
    End Property

    Public Property idParentesco() As Int16
        Get
            Return _idParentesco
        End Get
        Set(ByVal value As Int16)
            _idParentesco = value
        End Set
    End Property

    Public Property idParentescoCampania() As Int16
        Get
            Return _idParentescoCampania
        End Get
        Set(ByVal value As Int16)
            _idParentescoCampania = value
        End Set
    End Property

End Class
