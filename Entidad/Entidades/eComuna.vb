Public Class eComuna

    Private _idComuna As Int32
    Private _nombreComuna As String
    Private _idCiudad As Int32
    Private _activo As Boolean

    Public Property idComuna() As Int32
        Get
            Return _idComuna
        End Get
        Set(ByVal value As Int32)
            _idComuna = value
        End Set
    End Property

    Public Property nombreComuna() As String
        Get
            Return _nombreComuna
        End Get
        Set(ByVal value As String)
            _nombreComuna = value
        End Set
    End Property

    Public Property idCiudad() As Int32
        Get
            Return _idCiudad
        End Get
        Set(ByVal value As Int32)
            _idCiudad = value
        End Set
    End Property

    Public Property activo() As Boolean
        Get
            Return _activo
        End Get
        Set(ByVal value As Boolean)
            _activo = value
        End Set
    End Property

End Class
