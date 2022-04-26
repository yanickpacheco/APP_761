Public Class eCliente

    Private _C_ID As Int32
    Private _C_Rut As String
    Private _C_Dv As String
    Private _C_Nombre As String
    Private _C_Paterno As String
    Private _C_Materno As String
    Private _C_Fecha_Nacimiento As String
    Private _C_Sexo As String
    Private _C_Edad As String
    Private _C_Email As String
    Private _C_Direccion As String
    Private _C_IdComuna As String
    Private _C_IdCiudad As String
    Private _C_Comuna As String
    Private _C_Ciudad As String
    Private _C_Referencia As String
    Private _C_Fono1 As String
    Private _C_Telefono1 As String
    Private _C_Fono2 As String
    Private _C_Telefono2 As String
    Private _C_Fono3 As String
    Private _C_Telefono3 As String
    Private _C_Fono4 As String
    Private _C_Telefono4 As String
    Private _C_Fono5 As String
    Private _C_Telefono5 As String
    Private _C_Fono6 As String
    Private _C_Telefono6 As String
    Private _C_Fono7 As String
    Private _C_Telefono7 As String
    Private _C_Fono_Alternativo As String
    Private _C_Intentos_Max As Integer
    Private _C_TipoBase As Integer

    'CAMPOS EXTRAS


    Private _Extra1 As String
    Private _Extra2 As String
    Private _Extra3 As String
    Private _Extra4 As String
    Private _Extra5 As String
    Private _Extra6 As String
    Private _Extra7 As String
    Private _Extra8 As String
    Private _Extra9 As String
    Private _Extra10 As String
    Private _Extra11 As String
    Private _Extra12 As String
    Private _Extra13 As String
    Private _Extra14 As String
    Private _Extra15 As String
    Private _Extra16 As String
    Private _Extra17 As String
    Private _Extra18 As String
    Private _Extra19 As String
    Private _Extra20 As String
    Private _Extra21 As String
    Private _Extra22 As String
    Private _Extra23 As String
    Private _Extra24 As String
    Private _Extra25 As String
    Private _Extra26 As String
    Private _Extra27 As String
    Private _Extra28 As String
    Private _Extra29 As String
    Private _Extra30 As String
    Private _Extra31 As String
    Private _Extra32 As String
    Private _Extra33 As String
    Private _Extra34 As String
    Private _Extra35 As String
    Private _Extra36 As String
    Private _Extra37 As String
    Private _Extra38 As String
    Private _Extra39 As String
    Private _Extra40 As String

    Public Sub New(c_ID As Integer, c_Rut As String, c_Dv As String, c_Nombre As String, c_Paterno As String, c_Materno As String, c_Fecha_Nacimiento As String, c_Sexo As String, c_Edad As String, c_Email As String, c_Direccion As String, c_Comuna As String, c_Ciudad As String, c_Referencia As String, c_Fono1 As String, c_Telefono1 As String, c_Fono2 As String, c_Telefono2 As String, c_Fono3 As String, c_Telefono3 As String, c_Fono4 As String, c_Telefono4 As String, c_Fono5 As String, c_Telefono5 As String, c_Fono6 As String, c_Telefono6 As String, c_Fono7 As String, c_Telefono7 As String, c_Fono_Alternativo As String, extra1 As String, extra2 As String, extra3 As String, extra4 As String, extra5 As String, extra6 As String, extra7 As String, extra8 As String, extra9 As String, extra10 As String, extra11 As String, extra12 As String, extra13 As String, extra14 As String, extra15 As String, extra16 As String, extra17 As String, extra18 As String, extra19 As String, extra20 As String, extra21 As String, extra22 As String, extra23 As String, extra24 As String, extra25 As String, extra26 As String, extra27 As String, extra28 As String, extra29 As String, extra30 As String, extra31 As String, extra32 As String, extra33 As String, extra34 As String, extra35 As String, extra36 As String, extra37 As String, extra38 As String, extra39 As String, extra40 As String, c_IdComuna As String, c_IdCiudad As String, C_Intentos_Max As Integer, C_TipoBase As Integer)
        _C_ID = c_ID
        _C_Rut = c_Rut
        _C_Dv = c_Dv
        _C_Nombre = c_Nombre
        _C_Paterno = c_Paterno
        _C_Materno = c_Materno
        _C_Fecha_Nacimiento = c_Fecha_Nacimiento
        _C_Sexo = c_Sexo
        _C_Edad = c_Edad
        _C_Email = c_Email
        _C_Direccion = c_Direccion
        _C_IdComuna = c_IdComuna
        _C_IdCiudad = c_IdCiudad
        _C_Comuna = c_Comuna
        _C_Ciudad = c_Ciudad
        _C_Referencia = c_Referencia
        _C_Fono1 = c_Fono1
        _C_Telefono1 = c_Telefono1
        _C_Fono2 = c_Fono2
        _C_Telefono2 = c_Telefono2
        _C_Fono3 = c_Fono3
        _C_Telefono3 = c_Telefono3
        _C_Fono4 = c_Fono4
        _C_Telefono4 = c_Telefono4
        _C_Fono5 = c_Fono5
        _C_Telefono5 = c_Telefono5
        _C_Fono6 = c_Fono6
        _C_Telefono6 = c_Telefono6
        _C_Fono7 = c_Fono7
        _C_Telefono7 = c_Telefono7
        _C_Fono_Alternativo = c_Fono_Alternativo
        _Extra1 = extra1
        _Extra2 = extra2
        _Extra3 = extra3
        _Extra4 = extra4
        _Extra5 = extra5
        _Extra6 = extra6
        _Extra7 = extra7
        _Extra8 = extra8
        _Extra9 = extra9
        _Extra10 = extra10
        _Extra11 = extra11
        _Extra12 = extra12
        _Extra13 = extra13
        _Extra14 = extra14
        _Extra15 = extra15
        _Extra16 = extra16
        _Extra17 = extra17
        _Extra18 = extra18
        _Extra19 = extra19
        _Extra20 = extra20
        _Extra21 = extra21
        _Extra22 = extra22
        _Extra23 = extra23
        _Extra24 = extra24
        _Extra25 = extra25
        _Extra26 = extra26
        _Extra27 = extra27
        _Extra28 = extra28
        _Extra29 = extra29
        _Extra30 = extra30
        _Extra31 = extra31
        _Extra32 = extra32
        _Extra33 = extra33
        _Extra34 = extra34
        _Extra35 = extra35
        _Extra36 = extra36
        _Extra37 = extra37
        _Extra38 = extra38
        _Extra39 = extra39
        _Extra40 = extra40
        _C_Intentos_Max = C_Intentos_Max
        _C_TipoBase = C_TipoBase
    End Sub

    Public Sub New()
        _C_ID = 0
        _C_Rut = ""
        _C_Dv = ""
        _C_Nombre = ""
        _C_Paterno = ""
        _C_Materno = ""
        _C_Fecha_Nacimiento = ""
        _C_Sexo = ""
        _C_Edad = ""
        _C_Email = ""
        _C_Direccion = ""
        _C_IdComuna = ""
        _C_IdCiudad = ""
        _C_Comuna = ""
        _C_Ciudad = ""
        _C_Referencia = ""
        _C_Fono1 = ""
        _C_Telefono1 = ""
        _C_Fono2 = ""
        _C_Telefono2 = ""
        _C_Fono3 = ""
        _C_Telefono3 = ""
        _C_Fono4 = ""
        _C_Telefono4 = ""
        _C_Fono5 = ""
        _C_Telefono5 = ""
        _C_Fono6 = ""
        _C_Telefono6 = ""
        _C_Fono7 = ""
        _C_Telefono7 = ""
        _C_Fono_Alternativo = ""
        _Extra1 = ""
        _Extra2 = ""
        _Extra3 = ""
        _Extra4 = ""
        _Extra5 = ""
        _Extra6 = ""
        _Extra7 = ""
        _Extra8 = ""
        _Extra9 = ""
        _Extra10 = ""
        _Extra11 = ""
        _Extra12 = ""
        _Extra13 = ""
        _Extra14 = ""
        _Extra15 = ""
        _Extra16 = ""
        _Extra17 = ""
        _Extra18 = ""
        _Extra19 = ""
        _Extra20 = ""
        _Extra21 = ""
        _Extra22 = ""
        _Extra23 = ""
        _Extra24 = ""
        _Extra25 = ""
        _Extra26 = ""
        _Extra27 = ""
        _Extra28 = ""
        _Extra29 = ""
        _Extra30 = ""
        _Extra31 = ""
        _Extra32 = ""
        _Extra33 = ""
        _Extra34 = ""
        _Extra35 = ""
        _Extra36 = ""
        _Extra37 = ""
        _Extra38 = ""
        _Extra39 = ""
        _Extra40 = ""
        _C_Intentos_Max = 0
        _C_TipoBase = 0

    End Sub

    Public Property C_Id() As Integer
        Get
            Return _C_ID
        End Get
        Set(ByVal value As Integer)
            _C_ID = value
        End Set
    End Property

    Public Property C_Rut() As String
        Get
            Return _C_Rut
        End Get
        Set(ByVal value As String)
            _C_Rut = value
        End Set
    End Property

    Public Property C_Dv() As String
        Get
            Return _C_Dv
        End Get
        Set(ByVal value As String)
            _C_Dv = value
        End Set
    End Property

    Public Property C_Nombre() As String
        Get
            Return _C_Nombre
        End Get
        Set(ByVal value As String)
            _C_Nombre = value
        End Set
    End Property

    Public Property C_Paterno() As String
        Get
            Return _C_Paterno
        End Get
        Set(ByVal value As String)
            _C_Paterno = value
        End Set
    End Property

    Public Property C_Materno() As String
        Get
            Return _C_Materno
        End Get
        Set(ByVal value As String)
            _C_Materno = value
        End Set
    End Property

    Public Property C_Fecha_Nacimiento() As String
        Get
            Return _C_Fecha_Nacimiento
        End Get
        Set(ByVal value As String)
            _C_Fecha_Nacimiento = value
        End Set
    End Property

    Public Property C_Sexo() As String
        Get
            Return _C_Sexo
        End Get
        Set(ByVal value As String)
            _C_Sexo = value
        End Set
    End Property

    Public Property C_Edad() As String
        Get
            Return _C_Edad
        End Get
        Set(ByVal value As String)
            _C_Edad = value
        End Set
    End Property

    Public Property C_Email() As String
        Get
            Return _C_Email
        End Get
        Set(ByVal value As String)
            _C_Email = value
        End Set
    End Property

    Public Property C_Direccion() As String
        Get
            Return _C_Direccion
        End Get
        Set(ByVal value As String)
            _C_Direccion = value
        End Set
    End Property

    Public Property C_Comuna() As String
        Get
            Return _C_Comuna
        End Get
        Set(ByVal value As String)
            _C_Comuna = value
        End Set
    End Property

    Public Property C_Ciudad() As String
        Get
            Return _C_Ciudad
        End Get
        Set(ByVal value As String)
            _C_Ciudad = value
        End Set
    End Property

    Public Property C_Referencia() As String
        Get
            Return _C_Referencia
        End Get
        Set(ByVal value As String)
            _C_Referencia = value
        End Set
    End Property

    Public Property C_Fono1() As String
        Get
            Return _C_Fono1
        End Get
        Set(ByVal value As String)
            _C_Fono1 = value
        End Set
    End Property

    Public Property C_Telefono1() As String
        Get
            Return _C_Telefono1
        End Get
        Set(ByVal value As String)
            _C_Telefono1 = value
        End Set
    End Property

    Public Property C_Fono2() As String
        Get
            Return _C_Fono2
        End Get
        Set(ByVal value As String)
            _C_Fono2 = value
        End Set
    End Property

    Public Property C_Telefono2() As String
        Get
            Return _C_Telefono2
        End Get
        Set(ByVal value As String)
            _C_Telefono2 = value
        End Set
    End Property

    Public Property C_Fono3() As String
        Get
            Return _C_Fono3
        End Get
        Set(ByVal value As String)
            _C_Fono3 = value
        End Set
    End Property

    Public Property C_Telefono3() As String
        Get
            Return _C_Telefono3
        End Get
        Set(ByVal value As String)
            _C_Telefono3 = value
        End Set
    End Property

    Public Property C_Fono4() As String
        Get
            Return _C_Fono4
        End Get
        Set(ByVal value As String)
            _C_Fono4 = value
        End Set
    End Property

    Public Property C_Telefono4() As String
        Get
            Return _C_Telefono4
        End Get
        Set(ByVal value As String)
            _C_Telefono4 = value
        End Set
    End Property

    Public Property C_Fono5() As String
        Get
            Return _C_Fono5
        End Get
        Set(ByVal value As String)
            _C_Fono5 = value
        End Set
    End Property

    Public Property C_Telefono5() As String
        Get
            Return _C_Telefono5
        End Get
        Set(ByVal value As String)
            _C_Telefono5 = value
        End Set
    End Property

    Public Property C_Fono6() As String
        Get
            Return _C_Fono6
        End Get
        Set(ByVal value As String)
            _C_Fono6 = value
        End Set
    End Property

    Public Property C_Telefono6() As String
        Get
            Return _C_Telefono6
        End Get
        Set(ByVal value As String)
            _C_Telefono6 = value
        End Set
    End Property

    Public Property C_Fono7() As String
        Get
            Return _C_Fono7
        End Get
        Set(ByVal value As String)
            _C_Fono7 = value
        End Set
    End Property

    Public Property C_Telefono7() As String
        Get
            Return _C_Telefono7
        End Get
        Set(ByVal value As String)
            _C_Telefono7 = value
        End Set
    End Property

    Public Property C_Fono_Alternativo() As String
        Get
            Return _C_Fono_Alternativo
        End Get
        Set(ByVal value As String)
            _C_Fono_Alternativo = value
        End Set
    End Property


    'capos extras

    Public Property Extra1() As String
        Get
            Return _Extra1
        End Get
        Set(ByVal value As String)
            _Extra1 = value
        End Set
    End Property

    Public Property Extra2() As String
        Get
            Return _Extra2
        End Get
        Set(ByVal value As String)
            _Extra2 = value
        End Set
    End Property

    Public Property Extra3() As String
        Get
            Return _Extra3
        End Get
        Set(ByVal value As String)
            _Extra3 = value
        End Set
    End Property

    Public Property Extra4() As String
        Get
            Return _Extra4
        End Get
        Set(ByVal value As String)
            _Extra4 = value
        End Set
    End Property

    Public Property Extra5() As String
        Get
            Return _Extra5
        End Get
        Set(ByVal value As String)
            _Extra5 = value
        End Set
    End Property

    Public Property Extra6() As String
        Get
            Return _Extra6
        End Get
        Set(ByVal value As String)
            _Extra6 = value
        End Set
    End Property

    Public Property Extra7() As String
        Get
            Return _Extra7
        End Get
        Set(ByVal value As String)
            _Extra7 = value
        End Set
    End Property

    Public Property Extra8() As String
        Get
            Return _Extra8
        End Get
        Set(ByVal value As String)
            _Extra8 = value
        End Set
    End Property

    Public Property Extra9() As String
        Get
            Return _Extra9
        End Get
        Set(ByVal value As String)
            _Extra9 = value
        End Set
    End Property

    Public Property Extra10() As String
        Get
            Return _Extra10
        End Get
        Set(ByVal value As String)
            _Extra10 = value
        End Set
    End Property

    Public Property Extra11() As String
        Get
            Return _Extra11
        End Get
        Set(ByVal value As String)
            _Extra11 = value
        End Set
    End Property

    Public Property Extra12() As String
        Get
            Return _Extra12
        End Get
        Set(ByVal value As String)
            _Extra12 = value
        End Set
    End Property

    Public Property Extra13() As String
        Get
            Return _Extra13
        End Get
        Set(ByVal value As String)
            _Extra13 = value
        End Set
    End Property

    Public Property Extra14() As String
        Get
            Return _Extra14
        End Get
        Set(ByVal value As String)
            _Extra14 = value
        End Set
    End Property

    Public Property Extra15() As String
        Get
            Return _Extra15
        End Get
        Set(ByVal value As String)
            _Extra15 = value
        End Set
    End Property

    Public Property Extra16() As String
        Get
            Return _Extra16
        End Get
        Set(ByVal value As String)
            _Extra16 = value
        End Set
    End Property

    Public Property Extra17() As String
        Get
            Return _Extra17
        End Get
        Set(ByVal value As String)
            _Extra17 = value
        End Set
    End Property

    Public Property Extra18() As String
        Get
            Return _Extra18
        End Get
        Set(ByVal value As String)
            _Extra18 = value
        End Set
    End Property

    Public Property Extra19() As String
        Get
            Return _Extra19
        End Get
        Set(ByVal value As String)
            _Extra19 = value
        End Set
    End Property

    Public Property Extra20() As String
        Get
            Return _Extra20
        End Get
        Set(ByVal value As String)
            _Extra20 = value
        End Set
    End Property

    Public Property Extra21() As String
        Get
            Return _Extra21
        End Get
        Set(ByVal value As String)
            _Extra21 = value
        End Set
    End Property

    Public Property Extra22() As String
        Get
            Return _Extra22
        End Get
        Set(ByVal value As String)
            _Extra22 = value
        End Set
    End Property

    Public Property Extra23() As String
        Get
            Return _Extra23
        End Get
        Set(ByVal value As String)
            _Extra23 = value
        End Set
    End Property

    Public Property Extra24() As String
        Get
            Return _Extra24
        End Get
        Set(ByVal value As String)
            _Extra24 = value
        End Set
    End Property

    Public Property Extra25() As String
        Get
            Return _Extra25
        End Get
        Set(ByVal value As String)
            _Extra25 = value
        End Set
    End Property

    Public Property Extra26() As String
        Get
            Return _Extra26
        End Get
        Set(ByVal value As String)
            _Extra26 = value
        End Set
    End Property

    Public Property Extra27() As String
        Get
            Return _Extra27
        End Get
        Set(ByVal value As String)
            _Extra27 = value
        End Set
    End Property

    Public Property Extra28() As String
        Get
            Return _Extra28
        End Get
        Set(ByVal value As String)
            _Extra28 = value
        End Set
    End Property

    Public Property Extra29() As String
        Get
            Return _Extra29
        End Get
        Set(ByVal value As String)
            _Extra29 = value
        End Set
    End Property

    Public Property Extra30() As String
        Get
            Return _Extra30
        End Get
        Set(ByVal value As String)
            _Extra30 = value
        End Set
    End Property

    Public Property Extra31() As String
        Get
            Return _Extra31
        End Get
        Set(ByVal value As String)
            _Extra31 = value
        End Set
    End Property

    Public Property Extra32() As String
        Get
            Return _Extra32
        End Get
        Set(ByVal value As String)
            _Extra32 = value
        End Set
    End Property

    Public Property Extra33() As String
        Get
            Return _Extra33
        End Get
        Set(ByVal value As String)
            _Extra33 = value
        End Set
    End Property

    Public Property Extra34() As String
        Get
            Return _Extra34
        End Get
        Set(ByVal value As String)
            _Extra34 = value
        End Set
    End Property

    Public Property Extra35() As String
        Get
            Return _Extra35
        End Get
        Set(ByVal value As String)
            _Extra35 = value
        End Set
    End Property

    Public Property Extra36() As String
        Get
            Return _Extra36
        End Get
        Set(ByVal value As String)
            _Extra36 = value
        End Set
    End Property

    Public Property Extra37() As String
        Get
            Return _Extra37
        End Get
        Set(ByVal value As String)
            _Extra37 = value
        End Set
    End Property

    Public Property Extra38() As String
        Get
            Return _Extra38
        End Get
        Set(ByVal value As String)
            _Extra38 = value
        End Set
    End Property

    Public Property Extra39() As String
        Get
            Return _Extra39
        End Get
        Set(ByVal value As String)
            _Extra39 = value
        End Set
    End Property

    Public Property Extra40() As String
        Get
            Return _Extra40
        End Get
        Set(ByVal value As String)
            _Extra40 = value
        End Set
    End Property

    Public Property C_IdComuna() As String
        Get
            Return _C_IdComuna
        End Get
        Set(ByVal value As String)
            _C_IdComuna = value
        End Set
    End Property


    Public Property C_IdCiudad() As String
        Get
            Return _C_IdCiudad
        End Get
        Set(ByVal value As String)
            _C_IdCiudad = value
        End Set
    End Property

    Public Property C_Intentos_Max() As Integer
        Get
            Return _C_Intentos_Max
        End Get
        Set(ByVal value As Integer)
            _C_Intentos_Max = value
        End Set
    End Property

    Public Property C_TipoBase() As Integer
        Get
            Return _C_TipoBase
        End Get
        Set(ByVal value As Integer)
            _C_TipoBase = value
        End Set
    End Property



End Class