Public Class eCiudad

    Private _idCiudad As Int32
    Private _idRegion As Int32
    Private _nombreCiudad As String
    Private _activo As Boolean

    Public Property idCiudad() As Int32
        Get
            Return _idCiudad
        End Get
        Set(ByVal value As Int32)
            _idCiudad = value
        End Set
    End Property

    Public Property idRegion() As Int32
        Get
            Return _idRegion
        End Get
        Set(ByVal value As Int32)
            _idRegion = value
        End Set
    End Property

    Public Property nombreCiudad() As String
        Get
            Return _nombreCiudad
        End Get
        Set(ByVal value As String)
            _nombreCiudad = value
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
