Public Class eParentesco
    Private _idParentesco As Int16
    Private _nombreParentesco As String

    Public Property nombreParentesco() As String
        Get
            Return _nombreParentesco
        End Get
        Set(ByVal value As String)
            _nombreParentesco = value
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

End Class
