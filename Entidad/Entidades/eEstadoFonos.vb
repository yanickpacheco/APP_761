Public Class eEstadoFonos

    Private _efId As Int16
    Private _efDescripcion As String
    Private _efVigente As UShort

    Public Property efVigente() As UShort
        Get
            Return _efVigente
        End Get
        Set(ByVal value As UShort)
            _efVigente = value
        End Set
    End Property

    Public Property efDescripcion() As String
        Get
            Return _efDescripcion
        End Get
        Set(ByVal value As String)
            _efDescripcion = value
        End Set
    End Property

    Public Property efId() As Int16
        Get
            Return _efId
        End Get
        Set(ByVal value As Int16)
            _efId = value
        End Set
    End Property


End Class
