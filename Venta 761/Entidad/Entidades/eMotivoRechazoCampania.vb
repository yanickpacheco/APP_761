Public Class eMotivoRechazoCampania

    Private _idMotivosObjecionCampania As Int32
    Public Property idMotivosObjecionCampania() As Int32
        Get
            Return _idMotivosObjecionCampania
        End Get
        Set(ByVal value As Int32)
            _idMotivosObjecionCampania = value
        End Set
    End Property

    Private _idMotivosObjecion As Int32
    Public Property idMotivosObjecion() As Int32
        Get
            Return _idMotivosObjecion
        End Get
        Set(ByVal value As Int32)
            _idMotivosObjecion = value
        End Set
    End Property

    Private _espCodigo As Int32
    Public Property espCodigo() As Int32
        Get
            Return _espCodigo
        End Get
        Set(ByVal value As Int32)
            _espCodigo = value
        End Set
    End Property

    Private _calcodigo As Int32
    Public Property calcodigo() As Int32
        Get
            Return _calcodigo
        End Get
        Set(ByVal value As Int32)
            _calcodigo = value
        End Set
    End Property

    Private _otro As Boolean
    Public Property otro() As Boolean
        Get
            Return _otro
        End Get
        Set(ByVal value As Boolean)
            _otro = value
        End Set
    End Property



End Class
