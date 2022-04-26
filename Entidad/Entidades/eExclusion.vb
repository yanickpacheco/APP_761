Public Class eExclusion

    Dim _idExclusion As Integer
    Dim _Exclusion As String

    Public Sub New()
        IdExclusion = 0
        Exclusion = ""
    End Sub
    Public Sub New(idExclusion As Integer, exclusion As String)
        _idExclusion = idExclusion
        _Exclusion = exclusion
    End Sub



    Public Property IdExclusion As Integer
        Get
            Return _idExclusion
        End Get
        Set(value As Integer)
            _idExclusion = value
        End Set
    End Property

    Public Property Exclusion As String
        Get
            Return _Exclusion
        End Get
        Set(value As String)
            _Exclusion = value
        End Set
    End Property

End Class
