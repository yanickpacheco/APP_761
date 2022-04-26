Public Class eGesRegrabacion
    Private _cli_id As String
    Private _cli_primotivo_cal As String
    Private _cli_dprisubmotivo_cal As String
    Private _cli_comuna As String
    Private _cli_fechaventa As String
    Private _cli_agente As String
    Private _cli_nombre As String
    Private _cli_rut As String


    Public Property cli_id() As String
        Get
            Return _cli_id
        End Get
        Set(ByVal value As String)
            _cli_id = value
        End Set
    End Property

    Public Property cli_comuna() As String
        Get
            Return _cli_comuna
        End Get
        Set(ByVal value As String)
            _cli_comuna = value
        End Set
    End Property
    Public Property cli_fechaventa() As String
        Get
            Return _cli_fechaventa
        End Get
        Set(ByVal value As String)
            _cli_fechaventa = value
        End Set
    End Property

    Public Property cli_agente() As String
        Get
            Return _cli_agente
        End Get
        Set(ByVal value As String)
            _cli_agente = value
        End Set
    End Property
    Public Property cli_nombre() As String
        Get
            Return _cli_nombre
        End Get
        Set(ByVal value As String)
            _cli_nombre = value
        End Set
    End Property

    Public Property cli_rut() As String
        Get
            Return _cli_rut
        End Get
        Set(ByVal value As String)
            _cli_rut = value
        End Set
    End Property


End Class
