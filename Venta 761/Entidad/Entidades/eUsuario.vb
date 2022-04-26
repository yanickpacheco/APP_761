Public Class eUsuario
    Private _cli_user As String
    Private _cli_perfil As String

    Public Property cli_user() As String
        Get
            Return _cli_user
        End Get
        Set(ByVal value As String)
            _cli_user = value
        End Set
    End Property

    Public Property cli_perfil() As String
        Get
            Return _cli_perfil
        End Get
        Set(ByVal value As String)
            _cli_perfil = value
        End Set
    End Property

End Class
