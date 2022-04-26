Public Class eTipoContrato

    Private _idTipoContrato As Int16
    Private _nombreTipoContrato As String
    Private _cantidadAdicionales As Int16
    Private _cantidadBeneficiarios As Int16
    Private _activo As Boolean
    Private _definido As Boolean

    Public Property definido() As Boolean
        Get
            Return _definido
        End Get
        Set(ByVal value As Boolean)
            _definido = value
        End Set
    End Property


    Public Property idTipoContrato() As Int16
        Get
            Return _idTipoContrato
        End Get
        Set(ByVal value As Int16)
            _idTipoContrato = value
        End Set
    End Property

    Public Property nombreTipoContrato() As String
        Get
            Return _nombreTipoContrato
        End Get
        Set(ByVal value As String)
            _nombreTipoContrato = value
        End Set
    End Property

    Public Property cantidadAdicionales() As Int16
        Get
            Return _cantidadAdicionales
        End Get
        Set(ByVal value As Int16)
            _cantidadAdicionales = value
        End Set
    End Property

    Public Property cantidadBeneficiarios() As Int16
        Get
            Return _cantidadBeneficiarios
        End Get
        Set(ByVal value As Int16)
            _cantidadBeneficiarios = value
        End Set
    End Property

    Public Property activo() As Boolean
        Get
            Return _activo
        End Get
        Set(ByVal value As Boolean)
            _activo = value
        End Set
    End Property


End Class
