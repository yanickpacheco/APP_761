Public Class eScript

    Private _idScript As Int64
    Private _contenidoScript As String
    Private _idTipoScript As Int16
    Private _idCRM As Int32

    Public Sub New()

        idScript = 0
        contenidoScript = ""
        idTipoScript = 0
        idCRM = 0

    End Sub

    Public Property idScript() As Int64
        Get
            Return _idScript
        End Get
        Set(ByVal value As Int64)
            _idScript = value
        End Set
    End Property

    Public Property contenidoScript() As String
        Get
            Return _contenidoScript
        End Get
        Set(ByVal value As String)
            _contenidoScript = value
        End Set
    End Property

    Public Property idTipoScript() As Int16
        Get
            Return _idTipoScript
        End Get
        Set(ByVal value As Int16)
            _idTipoScript = value
        End Set
    End Property

    Public Property idCRM() As Int32
        Get
            Return _idCRM
        End Get
        Set(ByVal value As Int32)
            _idCRM = value
        End Set
    End Property



End Class
