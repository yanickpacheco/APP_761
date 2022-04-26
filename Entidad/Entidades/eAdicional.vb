Public Class eAdicional
    Private _A_ID As Integer
    Private _C_ID As Integer
    Private _A_NRO As Integer
    Private _A_Rut As String
    Private _A_Dv As String
    Private _A_Nombre As String
    Private _A_Paterno As String
    Private _A_Materno As String
    Private _A_TipoBeneficiario As Int16
    Private _A_FechaNacimiento As String
    Private _A_Sexo As String
    Private _A_NombreBeneficiario As String
    Private _a_primaUf As String
    Private _a_primaPesos As Int64
    Private _idPlanAdic As Int64
    Private _a_salud As String
    Public Property a_primaPesos() As Int64
        Get
            Return _a_primaPesos
        End Get
        Set(ByVal value As Int64)
            _a_primaPesos = value
        End Set
    End Property

    Public Property A_NombreBeneficiario() As String
        Get
            Return _A_NombreBeneficiario
        End Get
        Set(ByVal value As String)
            _A_NombreBeneficiario = value
        End Set
    End Property
    Public Property a_salud() As String
        Get
            Return _a_salud
        End Get
        Set(ByVal value As String)
            _a_salud = value
        End Set
    End Property
    Public Property idPlanAdic() As Int64
        Get
            Return _idPlanAdic
        End Get
        Set(ByVal value As Int64)
            _idPlanAdic = value
        End Set
    End Property

    Public Property a_primaUf() As String
        Get
            Return _a_primaUf
        End Get
        Set(ByVal value As String)
            _a_primaUf = value
        End Set
    End Property

    Public Sub New(a_ID As String, c_ID As String, a_NRO As Integer, a_Rut As String, a_Dv As String, a_Nombre As String, a_Paterno As String, a_Materno As String, a_TipoBeneficiario As Int16, a_FechaNacimiento As String, a_Sexo As String, a_primaUf As String, idPlanAdic As Int64, a_salud As String, A_NombreBeneficiario As String, a_primaPesos As Int64)
        _A_ID = a_ID
        _C_ID = c_ID
        _A_NRO = a_NRO
        _A_Rut = a_Rut
        _A_Dv = a_Dv
        _A_Nombre = a_Nombre
        _A_Paterno = a_Paterno
        _A_Materno = a_Materno
        _A_TipoBeneficiario = a_TipoBeneficiario
        _A_FechaNacimiento = a_FechaNacimiento
        _A_Sexo = a_Sexo
        _a_primaUf = a_primaUf
        _a_primaPesos = a_primaPesos
        _idPlanAdic = idPlanAdic
        _a_salud = a_salud
        _A_NombreBeneficiario = A_NombreBeneficiario
    End Sub

    Public Sub New()
        _A_ID = 0
        _C_ID = 0
        _A_NRO = 0
        _A_Rut = ""
        _A_Dv = ""
        _A_Nombre = ""
        _A_Paterno = ""
        _A_Materno = ""
        _A_TipoBeneficiario = 0
        _A_FechaNacimiento = ""
        _A_Sexo = ""
        _a_primaUf = ""
        _a_primaPesos = 0
        _idPlanAdic = 0
        _a_salud = ""
        _A_NombreBeneficiario = ""
    End Sub

    Public Property A_ID As String
        Get
            Return _A_ID
        End Get
        Set(value As String)
            _A_ID = value
        End Set
    End Property

    Public Property C_ID As String
        Get
            Return _C_ID
        End Get
        Set(value As String)
            _C_ID = value
        End Set
    End Property

    Public Property A_NRO As Integer
        Get
            Return _A_NRO
        End Get
        Set(value As Integer)
            _A_NRO = value
        End Set
    End Property

    Public Property A_Rut As String
        Get
            Return _A_Rut
        End Get
        Set(value As String)
            _A_Rut = value
        End Set
    End Property

    Public Property A_Dv As String
        Get
            Return _A_Dv
        End Get
        Set(value As String)
            _A_Dv = value
        End Set
    End Property

    Public Property A_Nombre As String
        Get
            Return _A_Nombre
        End Get
        Set(value As String)
            _A_Nombre = value
        End Set
    End Property

    Public Property A_Paterno As String
        Get
            Return _A_Paterno
        End Get
        Set(value As String)
            _A_Paterno = value
        End Set
    End Property

    Public Property A_Materno As String
        Get
            Return _A_Materno
        End Get
        Set(value As String)
            _A_Materno = value
        End Set
    End Property

    Public Property A_TipoBeneficiario As Short
        Get
            Return _A_TipoBeneficiario
        End Get
        Set(value As Short)
            _A_TipoBeneficiario = value
        End Set
    End Property

    Public Property A_FechaNacimiento As String
        Get
            Return _A_FechaNacimiento
        End Get
        Set(value As String)
            _A_FechaNacimiento = value
        End Set
    End Property

    Public Property A_Sexo As String
        Get
            Return _A_Sexo
        End Get
        Set(value As String)
            _A_Sexo = value
        End Set
    End Property
End Class
