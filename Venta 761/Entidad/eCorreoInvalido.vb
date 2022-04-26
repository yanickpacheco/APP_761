Public Class eCorreoInvalido
    Private _idCorreoInvalido As Int64
    Private _correo As String
    Private _fechaIngreso As DateTime

    Public Property fechaIngreso() As DateTime
        Get
            Return _fechaIngreso
        End Get
        Set(ByVal value As DateTime)
            _fechaIngreso = value
        End Set
    End Property

    Public Property correo() As String
        Get
            Return _correo
        End Get
        Set(ByVal value As String)
            _correo = value
        End Set
    End Property

    Public Property idCorreoInvalido() As Int64
        Get
            Return _idCorreoInvalido
        End Get
        Set(ByVal value As Int64)
            _idCorreoInvalido = value
        End Set
    End Property
End Class
