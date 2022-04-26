Public Class eBeneficiario
    Private _B_ID As Integer
    Private _C_ID As Integer
    Private _B_Nro As String
    Private _B_Rut As String
    Private _B_Dv As String
    Private _B_Nombre As String
    Private _B_Paterno As String
    Private _B_Materno As String
    Private _B_TipoBeneficiario As Integer
    Private _B_FechaNacimiento As String
    Private _B_Porcentaje As String
    Private _B_Sexo As String
    Private _B_NOMBRE_PARENTESCO As String
    Public Property B_NOMBRE_PARENTESCO() As String
        Get
            Return _B_NOMBRE_PARENTESCO
        End Get
        Set(ByVal value As String)
            _B_NOMBRE_PARENTESCO = value
        End Set
    End Property

    Public Sub New(b_ID As Integer, c_ID As Integer, b_Nro As String, b_Rut As String, b_Dv As String, b_Nombre As String, b_Paterno As String, b_Materno As String, b_TipoBeneficiario As Integer, b_FechaNacimiento As String, b_Porcentaje As String, b_Sexo As String, B_NOMBRE_PARENTESCO As String)
        _B_ID = b_ID
        _C_ID = c_ID
        _B_Nro = b_Nro
        _B_Rut = b_Rut
        _B_Dv = b_Dv
        _B_Nombre = b_Nombre
        _B_Paterno = b_Paterno
        _B_Materno = b_Materno
        _B_TipoBeneficiario = b_TipoBeneficiario
        _B_FechaNacimiento = b_FechaNacimiento
        _B_Porcentaje = b_Porcentaje
        _B_Sexo = b_Sexo
        _B_NOMBRE_PARENTESCO = B_NOMBRE_PARENTESCO
    End Sub

    Public Sub New()
        _B_ID = 0
        _C_ID = 0
        _B_Nro = ""
        _B_Rut = ""
        _B_Dv = ""
        _B_Nombre = ""
        _B_Paterno = ""
        _B_Materno = ""
        _B_TipoBeneficiario = 0
        _B_FechaNacimiento = ""
        _B_Porcentaje = ""
        _B_Sexo = ""
        _B_NOMBRE_PARENTESCO = ""
    End Sub

    Public Property B_ID As Integer
        Get
            Return _B_ID
        End Get
        Set(value As Integer)
            _B_ID = value
        End Set
    End Property

    Public Property C_ID As Integer
        Get
            Return _C_ID
        End Get
        Set(value As Integer)
            _C_ID = value
        End Set
    End Property

    Public Property B_Nro As String
        Get
            Return _B_Nro
        End Get
        Set(value As String)
            _B_Nro = value
        End Set
    End Property

    Public Property B_Rut As String
        Get
            Return _B_Rut
        End Get
        Set(value As String)
            _B_Rut = value
        End Set
    End Property

    Public Property B_Dv As String
        Get
            Return _B_Dv
        End Get
        Set(value As String)
            _B_Dv = value
        End Set
    End Property

    Public Property B_Nombre As String
        Get
            Return _B_Nombre
        End Get
        Set(value As String)
            _B_Nombre = value
        End Set
    End Property

    Public Property B_Paterno As String
        Get
            Return _B_Paterno
        End Get
        Set(value As String)
            _B_Paterno = value
        End Set
    End Property

    Public Property B_Materno As String
        Get
            Return _B_Materno
        End Get
        Set(value As String)
            _B_Materno = value
        End Set
    End Property

    Public Property B_TipoBeneficiario As Integer
        Get
            Return _B_TipoBeneficiario
        End Get
        Set(value As Integer)
            _B_TipoBeneficiario = value
        End Set
    End Property

    Public Property B_FechaNacimiento As String
        Get
            Return _B_FechaNacimiento
        End Get
        Set(value As String)
            _B_FechaNacimiento = value
        End Set
    End Property

    Public Property B_Porcentaje As String
        Get
            Return _B_Porcentaje
        End Get
        Set(value As String)
            _B_Porcentaje = value
        End Set
    End Property

    Public Property B_Sexo As String
        Get
            Return _B_Sexo
        End Get
        Set(value As String)
            _B_Sexo = value
        End Set
    End Property
End Class
