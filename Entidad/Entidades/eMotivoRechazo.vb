Public Class eMotivoRechazo


    Private _idMotivosObjecion As Int32
    Private _nombreMotivosObjecion As String
    Private _activo As Boolean

    Public Sub New()

        idMotivosObjecion = 0
        nombreMotivosObjecion = ""
        activo = False

    End Sub

    Public Property idMotivosObjecion() As Int32
        Get
            Return _idMotivosObjecion
        End Get
        Set(ByVal value As Int32)
            _idMotivosObjecion = value
        End Set
    End Property


    Public Property nombreMotivosObjecion() As String
        Get
            Return _nombreMotivosObjecion
        End Get
        Set(ByVal value As String)
            _nombreMotivosObjecion = value
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
