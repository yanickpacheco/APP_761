Public Class clsAdicional
    Dim _nombre As String
    Dim _nombre2 As String
    Dim _paterno As String
    Dim _materno As String
    Dim _rut As String
    Dim _dv As String
    Dim _parentesco As String
    Dim _plan As String
    Dim _fechaNacimiento As String


    Public Sub New(ByVal nombre As String, ByVal nombre2 As String, ByVal paterno As String, ByVal materno As String, ByVal rut As String, ByVal dv As String, ByVal parentesco As String, ByVal plan As String, ByVal fechanacimiento As String)

        _nombre = nombre
        _nombre2 = nombre2
        _paterno = paterno
        _materno = materno
        _rut = rut
        _dv = dv
        _parentesco = parentesco
        _plan = plan
        _fechaNacimiento = fechanacimiento

    End Sub

    Public Sub New()
        _nombre = String.Empty
        _nombre2 = String.Empty
        _paterno = String.Empty
        _materno = String.Empty
        _rut = String.Empty
        _dv = String.Empty
        _parentesco = String.Empty
        _plan = String.Empty
        _fechaNacimiento = String.Empty
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

    Private _idPlan As Int64
    Public Property idPlan() As Int64
        Get
            Return _idPlan
        End Get
        Set(ByVal value As Int64)
            _idPlan = value
        End Set
    End Property


    Public Property plan() As String
        Get
            Return _plan
        End Get
        Set(ByVal value As String)
            _plan = value
        End Set
    End Property

    Public Property FechaNacimiento() As String
        Get
            Return _fechaNacimiento
        End Get
        Set(ByVal value As String)
            _fechaNacimiento = value
        End Set
    End Property

    Private _sexo As String
    Public Property sexo() As String
        Get
            Return _sexo
        End Get
        Set(ByVal value As String)
            _sexo = value
        End Set
    End Property



End Class
