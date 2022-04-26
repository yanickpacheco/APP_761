Public Class eGestion
    Private _G_ID As Integer
    Private _C_ID As Integer
    Private _G_Estado As String
    Private _G_Fecha As String
    Private _G_Hora As String
    Private _G_Fecha_Vta As String
    Private _G_Hora_Vta As String
    Private _G_IdLlamada As String
    Private _G_Intento As Integer
    Private _G_Ejecutivo As String
    Private _G_Ip_Ejecutivo As String
    Private _G_Fecha_Agen As String
    Private _G_Hora_Agen As String
    Private _G_Obs_Agen As String
    Private _G_Contacto As String
    Private _G_No_Contacto As String
    Private _G_Contacto_Con As String
    Private _G_Contacto_Flujo As String
    Private _G_Motivo_No_Interesa As String
    Private _G_Rut As String
    Private _G_Dv As String
    Private _G_Nombre As String
    Private _G_Paterno As String
    Private _G_Materno As String
    Private _G_Fecha_Nacimiento As String
    Private _G_Sexo As String
    Private _G_Email As String
    Private _G_Calle As String
    Private _G_Nro As String
    Private _G_Referencia As String
    Private _G_Comuna As String
    Private _G_Ciudad As String
    Private _G_IdComuna As String
    Private _G_IdCiudad As String
    Private _G_Fono_Contacto As String
    Private _G_TipoContrato As Integer
    Private _G_Plan As Integer
    Private _G_Prima_Uf As String
    Private _G_Prima_Pesos As String
    Private _G_Venta As Integer
    Private _G_Dato1 As String
    Private _G_Dato2 As String
    Private _G_Dato3 As String
    Private _G_Tiempo_PostView As String
    Private _G_BANCO As String
    Private _G_VENCIMIENTO_TARJETA As String
    Private _G_NUMERO_TARJETA As String
    Private _G_MEDIO_PAGO As String
    Private _G_TIPO_TARJETA As String
    Private _G_DIA_PAGO As String
    Public Property G_DIA_PAGO() As String
        Get
            Return _G_DIA_PAGO
        End Get
        Set(ByVal value As String)
            _G_DIA_PAGO = value
        End Set
    End Property
    Public Property G_TIPO_TARJETA() As String
        Get
            Return _G_TIPO_TARJETA
        End Get
        Set(ByVal value As String)
            _G_TIPO_TARJETA = value
        End Set
    End Property
    Public Property G_MEDIO_PAGO() As String
        Get
            Return _G_MEDIO_PAGO
        End Get
        Set(ByVal value As String)
            _G_MEDIO_PAGO = value
        End Set
    End Property
    Public Property G_NUMERO_TARJETA() As String
        Get
            Return _G_NUMERO_TARJETA
        End Get
        Set(ByVal value As String)
            _G_NUMERO_TARJETA = value
        End Set
    End Property
    Public Property G_VENCIMIENTO_TARJETA() As String
        Get
            Return _G_VENCIMIENTO_TARJETA
        End Get
        Set(ByVal value As String)
            _G_VENCIMIENTO_TARJETA = value
        End Set
    End Property
    Public Property G_BANCO() As String
        Get
            Return _G_BANCO
        End Get
        Set(ByVal value As String)
            _G_BANCO = value
        End Set
    End Property

    Public Sub New(g_ID As Integer, c_ID As Integer, g_Estado As String, g_Fecha As String, g_Hora As String, g_Fecha_Vta As String, g_Hora_Vta As String, G_IdLlamada As String, g_Intento As Integer, g_Ejecutivo As String, g_Ip_Ejecutivo As String, g_Fecha_Agen As String, g_Hora_Agen As String, g_Obs_Agen As String, g_Contacto As String, g_No_Contacto As String, g_Contacto_Con As String, g_Contacto_Flujo As String, g_Rut As String, g_Dv As String, g_Nombre As String, g_Paterno As String, g_Materno As String, g_Fecha_Nacimiento As String, g_Sexo As String, g_Email As String, g_Calle As String, g_Nro As String, g_Referencia As String, g_Comuna As String, g_Ciudad As String, g_IdComuna As String, g_IdCiudad As String, g_Fono_Contacto As String, g_TipoContrato As Integer, g_Plan As Integer, g_Prima_Uf As String, g_Prima_Pesos As String, g_Venta As String, G_Motivo_No_Interesa As String, G_Dato1 As String, G_Dato2 As String, G_Dato3 As String, G_Tiempo_PostView As String)
        Me.G_ID = g_ID
        Me.C_ID = c_ID
        Me.G_Estado = g_Estado
        Me.G_Fecha = g_Fecha
        Me.G_Hora = g_Hora
        Me.G_Fecha_Vta = g_Fecha_Vta
        Me.G_Hora_Vta = g_Hora_Vta
        Me.G_IdLlamada = G_IdLlamada
        Me.G_Intento = g_Intento
        Me.G_Ejecutivo = g_Ejecutivo
        Me.G_Ip_Ejecutivo = g_Ip_Ejecutivo
        Me.G_Fecha_Agen = g_Fecha_Agen
        Me.G_Hora_Agen = g_Hora_Agen
        Me.G_Obs_Agen = g_Obs_Agen
        Me.G_Contacto = g_Contacto
        Me.G_No_Contacto = g_No_Contacto
        Me.G_Contacto_Con = g_Contacto_Con
        Me.G_Contacto_Flujo = g_Contacto_Flujo
        Me._G_Motivo_No_Interesa = G_Motivo_No_Interesa
        Me.G_Rut = g_Rut
        Me.G_Dv = g_Dv
        Me.G_Nombre = g_Nombre
        Me.G_Paterno = g_Paterno
        Me.G_Materno = g_Materno
        Me.G_Fecha_Nacimiento = g_Fecha_Nacimiento
        Me.G_Sexo = g_Sexo
        Me.G_Email = g_Email
        Me.G_Calle = g_Calle
        Me.G_Nro = g_Nro
        Me.G_Referencia = g_Referencia
        Me.G_Comuna = g_Comuna
        Me.G_Ciudad = g_Ciudad
        Me.G_IdComuna = g_IdComuna
        Me.G_IdCiudad = g_IdCiudad
        Me.G_Fono_Contacto = g_Fono_Contacto
        Me.G_TipoContrato = g_TipoContrato
        Me.G_Plan = g_Plan
        Me.G_Prima_Uf = g_Prima_Uf
        Me.G_Prima_Pesos = g_Prima_Pesos
        Me.G_Venta = g_Venta
        Me._G_Dato1 = G_Dato1
        Me._G_Dato2 = G_Dato2
        Me._G_Dato3 = G_Dato3
        Me._G_Tiempo_PostView = G_Tiempo_PostView
    End Sub
    Public Sub New()
        Me.G_ID = 0
        Me.C_ID = 0
        Me.G_Estado = ""
        Me.G_Fecha = ""
        Me.G_Hora = ""
        Me.G_Fecha_Vta = ""
        Me.G_Hora_Vta = ""
        Me.G_IdLlamada = ""
        Me.G_Intento = 0
        Me.G_Ejecutivo = ""
        Me.G_Ip_Ejecutivo = ""
        Me.G_Fecha_Agen = ""
        Me.G_Hora_Agen = ""
        Me.G_Obs_Agen = ""
        Me.G_Contacto = ""
        Me.G_No_Contacto = ""
        Me.G_Contacto_Con = ""
        Me.G_Contacto_Flujo = ""
        Me._G_Motivo_No_Interesa = ""
        Me.G_Rut = ""
        Me.G_Dv = ""
        Me.G_Nombre = ""
        Me.G_Paterno = ""
        Me.G_Materno = ""
        Me.G_Fecha_Nacimiento = ""
        Me.G_Sexo = ""
        Me.G_Email = ""
        Me.G_Calle = ""
        Me.G_Nro = ""
        Me.G_Referencia = ""
        Me.G_Comuna = ""
        Me.G_Ciudad = ""
        Me.G_IdComuna = ""
        Me.G_IdCiudad = ""
        Me.G_Fono_Contacto = ""
        Me.G_TipoContrato = 0
        Me.G_Plan = 0
        Me.G_Prima_Uf = ""
        Me.G_Prima_Pesos = ""
        Me.G_Venta = 0
        Me.G_Dato1 = ""
        Me.G_Dato2 = ""
        Me.G_Dato3 = ""
        Me._G_Tiempo_PostView = ""
    End Sub
    Public Property G_ID As Integer
        Get
            Return _G_ID
        End Get
        Set(value As Integer)
            _G_ID = value
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

    Public Property G_Estado As String
        Get
            Return _G_Estado
        End Get
        Set(value As String)
            _G_Estado = value
        End Set
    End Property

    Public Property G_Fecha As String
        Get
            Return _G_Fecha
        End Get
        Set(value As String)
            _G_Fecha = value
        End Set
    End Property

    Public Property G_Hora As String
        Get
            Return _G_Hora
        End Get
        Set(value As String)
            _G_Hora = value
        End Set
    End Property

    Public Property G_Fecha_Vta As String
        Get
            Return _G_Fecha_Vta
        End Get
        Set(value As String)
            _G_Fecha_Vta = value
        End Set
    End Property

    Public Property G_Hora_Vta As String
        Get
            Return _G_Hora_Vta
        End Get
        Set(value As String)
            _G_Hora_Vta = value
        End Set
    End Property

    Public Property G_IdLlamada As String
        Get
            Return _G_IdLlamada
        End Get
        Set(value As String)
            _G_IdLlamada = value
        End Set
    End Property

    Public Property G_Intento As Integer
        Get
            Return _G_Intento
        End Get
        Set(value As Integer)
            _G_Intento = value
        End Set
    End Property

    Public Property G_Ejecutivo As String
        Get
            Return _G_Ejecutivo
        End Get
        Set(value As String)
            _G_Ejecutivo = value
        End Set
    End Property

    Public Property G_Ip_Ejecutivo As String
        Get
            Return _G_Ip_Ejecutivo
        End Get
        Set(value As String)
            _G_Ip_Ejecutivo = value
        End Set
    End Property

    Public Property G_Fecha_Agen As String
        Get
            Return _G_Fecha_Agen
        End Get
        Set(value As String)
            _G_Fecha_Agen = value
        End Set
    End Property

    Public Property G_Hora_Agen As String
        Get
            Return _G_Hora_Agen
        End Get
        Set(value As String)
            _G_Hora_Agen = value
        End Set
    End Property

    Public Property G_Obs_Agen As String
        Get
            Return _G_Obs_Agen
        End Get
        Set(value As String)
            _G_Obs_Agen = value
        End Set
    End Property

    Public Property G_Contacto As String
        Get
            Return _G_Contacto
        End Get
        Set(value As String)
            _G_Contacto = value
        End Set
    End Property

    Public Property G_No_Contacto As String
        Get
            Return _G_No_Contacto
        End Get
        Set(value As String)
            _G_No_Contacto = value
        End Set
    End Property

    Public Property G_Contacto_Con As String
        Get
            Return _G_Contacto_Con
        End Get
        Set(value As String)
            _G_Contacto_Con = value
        End Set
    End Property

    Public Property G_Contacto_Flujo As String
        Get
            Return _G_Contacto_Flujo
        End Get
        Set(value As String)
            _G_Contacto_Flujo = value
        End Set
    End Property

    Public Property G_Motivo_No_Interesa As String
        Get
            Return _G_Motivo_No_Interesa
        End Get
        Set(value As String)
            _G_Motivo_No_Interesa = value
        End Set
    End Property

    Public Property G_Rut As String
        Get
            Return _G_Rut
        End Get
        Set(value As String)
            _G_Rut = value
        End Set
    End Property

    Public Property G_Dv As String
        Get
            Return _G_Dv
        End Get
        Set(value As String)
            _G_Dv = value
        End Set
    End Property

    Public Property G_Nombre As String
        Get
            Return _G_Nombre
        End Get
        Set(value As String)
            _G_Nombre = value
        End Set
    End Property

    Public Property G_Paterno As String
        Get
            Return _G_Paterno
        End Get
        Set(value As String)
            _G_Paterno = value
        End Set
    End Property

    Public Property G_Materno As String
        Get
            Return _G_Materno
        End Get
        Set(value As String)
            _G_Materno = value
        End Set
    End Property

    Public Property G_Fecha_Nacimiento As String
        Get
            Return _G_Fecha_Nacimiento
        End Get
        Set(value As String)
            _G_Fecha_Nacimiento = value
        End Set
    End Property

    Public Property G_Sexo As String
        Get
            Return _G_Sexo
        End Get
        Set(value As String)
            _G_Sexo = value
        End Set
    End Property

    Public Property G_Email As String
        Get
            Return _G_Email
        End Get
        Set(value As String)
            _G_Email = value
        End Set
    End Property

    Public Property G_Calle As String
        Get
            Return _G_Calle
        End Get
        Set(value As String)
            _G_Calle = value
        End Set
    End Property

    Public Property G_Nro As String
        Get
            Return _G_Nro
        End Get
        Set(value As String)
            _G_Nro = value
        End Set
    End Property

    Public Property G_Referencia As String
        Get
            Return _G_Referencia
        End Get
        Set(value As String)
            _G_Referencia = value
        End Set
    End Property

    Public Property G_Comuna As String
        Get
            Return _G_Comuna
        End Get
        Set(value As String)
            _G_Comuna = value
        End Set
    End Property

    Public Property G_Ciudad As String
        Get
            Return _G_Ciudad
        End Get
        Set(value As String)
            _G_Ciudad = value
        End Set
    End Property

    Public Property G_IdComuna As String
        Get
            Return _G_IdComuna
        End Get
        Set(value As String)
            _G_IdComuna = value
        End Set
    End Property

    Public Property G_IdCiudad As String
        Get
            Return _G_IdCiudad
        End Get
        Set(value As String)
            _G_IdCiudad = value
        End Set
    End Property

    Public Property G_Fono_Contacto As String
        Get
            Return _G_Fono_Contacto
        End Get
        Set(value As String)
            _G_Fono_Contacto = value
        End Set
    End Property

    Public Property G_TipoContrato As Integer
        Get
            Return _G_TipoContrato
        End Get
        Set(value As Integer)
            _G_TipoContrato = value
        End Set
    End Property
    Public Property G_Plan As Integer
        Get
            Return _G_Plan
        End Get
        Set(value As Integer)
            _G_Plan = value
        End Set
    End Property

    Public Property G_Prima_Uf As String
        Get
            Return _G_Prima_Uf
        End Get
        Set(value As String)
            _G_Prima_Uf = value
        End Set
    End Property

    Public Property G_Prima_Pesos As String
        Get
            Return _G_Prima_Pesos
        End Get
        Set(value As String)
            _G_Prima_Pesos = value
        End Set
    End Property

    Public Property G_Venta As Integer
        Get
            Return _G_Venta
        End Get
        Set(value As Integer)
            _G_Venta = value
        End Set
    End Property

    Public Property G_Dato1 As String
        Get
            Return _G_Dato1
        End Get
        Set(value As String)
            _G_Dato1 = value
        End Set
    End Property

    Public Property G_Dato2 As String
        Get
            Return _G_Dato2
        End Get
        Set(value As String)
            _G_Dato2 = value
        End Set
    End Property

    Public Property G_Dato3 As String
        Get
            Return _G_Dato3
        End Get
        Set(value As String)
            _G_Dato3 = value
        End Set
    End Property

    Public Property G_Tiempo_PostView As String
        Get
            Return _G_Tiempo_PostView
        End Get
        Set(value As String)
            _G_Tiempo_PostView = value
        End Set
    End Property
End Class
