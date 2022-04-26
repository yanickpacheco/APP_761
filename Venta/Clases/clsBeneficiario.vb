Public Class clsBeneficiario
    Dim _nombre As String
    Dim _nombre2 As String
    Dim _paterno As String
    Dim _materno As String
    Dim _rut As String
    Dim _dv As String
    Dim _parentesco As String
    Dim _porcentaje As String
    Dim _idParentesco As Int16

    Public Sub New(ByVal nombre As String, ByVal nombre2 As String, ByVal paterno As String, ByVal materno As String, ByVal rut As String, ByVal dv As String, ByVal parentesco As String, ByVal porcentaje As String, ByVal idParentesco As Int16)
        _nombre = nombre
        _nombre2 = nombre2
        _paterno = paterno
        _materno = materno
        _rut = rut
        _dv = dv
        _parentesco = parentesco
        _porcentaje = porcentaje
        _idParentesco = idParentesco
    End Sub

    Public Sub New()
        _nombre = String.Empty
        _nombre2 = String.Empty
        _paterno = String.Empty
        _materno = String.Empty
        _rut = String.Empty
        _dv = String.Empty
        _parentesco = String.Empty
        _porcentaje = String.Empty
        _idParentesco = 0
    End Sub

    Public Property Nombre() As String
        Get
            Return _nombre
        End Get
        Set(ByVal value As String)
            _nombre = value
        End Set
    End Property

    Public Property Nombre2() As String
        Get
            Return _nombre2
        End Get
        Set(ByVal value As String)
            _nombre2 = value
        End Set
    End Property

    Public Property Paterno() As String
        Get
            Return _paterno
        End Get
        Set(ByVal value As String)
            _paterno = value
        End Set
    End Property

    Public Property Materno() As String
        Get
            Return _materno
        End Get
        Set(ByVal value As String)
            _materno = value
        End Set
    End Property

    Public Property Rut() As String
        Get
            Return _rut
        End Get
        Set(ByVal value As String)
            _rut = value
        End Set
    End Property

    Public Property Dv() As String
        Get
            Return _dv
        End Get
        Set(ByVal value As String)
            _dv = value
        End Set
    End Property

    Public Property Parentesco() As String
        Get
            Return _parentesco
        End Get
        Set(ByVal value As String)
            _parentesco = value
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
    Public Property Porcentaje() As String
        Get
            Return _porcentaje
        End Get
        Set(ByVal value As String)
            _porcentaje = value
        End Set
    End Property

End Class
