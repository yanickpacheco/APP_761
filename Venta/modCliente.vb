Option Strict Off
Option Explicit On

Imports Entidad

Module modCliente
    Public CLIENTE As New eCliente
    Public GESTION As New eGestion
    Public RECUPERACION As New eRecuperacion

    ' variable global que guardara los datos del cliente y su gestion hasta el final
    Public Function inicializarCliente(ByRef row As DataRow) As eCliente

        With CLIENTE
            '----------------------------
            .C_Id = row("C_ID") & ""
            .C_Rut = row("C_RUT") & ""
            .C_Dv = row("C_DV") & ""
            .C_Edad = row("C_EDAD") & ""
            .C_Fecha_Nacimiento = row("C_FECHANACIMIENTO") & ""
            .C_Nombre = row("C_NOMBRE") & ""
            .C_Paterno = row("C_PATERNO") & ""
            .C_Materno = row("C_MATERNO") & ""
            .C_Sexo = row("C_SEXO") & ""
            .C_Direccion = row("C_DIRECCION") & ""
            .C_Comuna = row("C_COMUNA") & ""
            .C_Ciudad = row("C_CIUDAD") & ""
            .C_Referencia = row("C_REFERENCIA") & ""
            .C_Email = row("C_EMAIL") & ""
            .C_Fono1 = row("C_FONO1") & ""
            .C_Telefono1 = row("C_TELEFONO1") & ""
            .C_Fono2 = row("C_FONO2") & ""
            .C_Telefono2 = row("C_TELEFONO2") & ""
            .C_Fono3 = row("C_FONO3") & ""
            .C_Telefono3 = row("C_TELEFONO3") & ""
            .C_Fono4 = row("C_FONO4") & ""
            .C_Telefono4 = row("C_TELEFONO4") & ""
            .C_Fono5 = row("C_FONO5") & ""
            .C_Telefono5 = row("C_TELEFONO5") & ""
            .C_Fono6 = row("C_FONO6") & ""
            .C_Telefono6 = row("C_TELEFONO6") & ""
            .C_Fono_Alternativo = row("C_FONOALTERNATIVO") & ""
            .C_Intentos_Max = row("C_INTENTOS_MAX")
            .C_TipoBase = row("C_TIPOBASE") & ""



            'If IsDBNull(row("CLI_INTENTOS_MAX")) Then
            '    .cli_intentos_max = "5"
            'Else
            '    .cli_intentos_max = Trim(row("CLI_INTENTOS_MAX"))
            'End If





            .Extra1 = row("C_extra1") & ""
            .Extra2 = row("C_extra2") & ""
            .Extra3 = row("C_extra3") & ""
            .Extra4 = row("C_extra4") & ""
            .Extra5 = row("C_extra5") & ""
            .Extra6 = row("C_extra6") & ""
            .Extra7 = row("C_extra7") & ""
            .Extra8 = row("C_extra8") & ""
            .Extra9 = row("C_extra9") & ""
            .Extra10 = row("C_extra10") & ""
            .Extra11 = row("C_extra11") & ""
            .Extra12 = row("C_extra12") & ""
            .Extra13 = row("C_extra13") & ""
            .Extra14 = row("C_extra14") & ""
            .Extra15 = row("C_extra15") & ""
            .Extra16 = row("C_extra16") & ""
            .Extra17 = row("C_extra17") & ""
            .Extra18 = row("C_extra18") & ""
            .Extra19 = row("C_extra19") & ""
            .Extra20 = row("C_extra20") & ""
            .Extra21 = row("C_extra21") & ""
            .Extra22 = row("C_extra22") & ""
            .Extra23 = row("C_extra23") & ""
            .Extra24 = row("C_extra24") & ""
            .Extra25 = row("C_extra25") & ""
            .Extra26 = row("C_extra26") & ""
            .Extra27 = row("C_extra27") & ""
            .Extra28 = row("C_extra28") & ""
            .Extra29 = row("C_extra29") & ""
            .Extra30 = row("C_extra30") & ""
            .Extra31 = row("C_extra31") & ""
            .Extra32 = row("C_extra32") & ""
            .Extra33 = row("C_extra33") & ""
            .Extra34 = row("C_extra34") & ""
            .Extra35 = row("C_extra35") & ""
            .Extra36 = row("C_extra36") & ""
            .Extra37 = row("C_extra37") & ""
            .Extra38 = row("C_extra38") & ""
            .Extra39 = row("C_extra39") & ""
            .Extra40 = row("C_extra40") & ""
        End With

        With GESTION
            .G_Intento = row("G_Intento") & ""
            .G_Obs_Agen = row("G_OBS_AGEN") & ""

            If IsDBNull(row("G_Intento")) Then
                .G_Intento = 1
            Else
                .G_Intento = row("G_Intento") + 1
            End If

        End With

        inicializarCliente = CLIENTE
        If db_central = 4 Then
            vpPosicion.LLAMANDO = False
        End If
    End Function


    Public Function inicializarGestionRecuperacion(ByRef row As DataRow) As eGestion

        With GESTION
            '----------------------------
            .C_ID = row("C_ID") & ""
            .G_Rut = row("G_RUT") & ""
            .G_Dv = row("G_DV") & ""
            .G_Fecha_Nacimiento = row("G_FECHANACIMIENTO") & ""
            .G_Nombre = row("G_NOMBRE") & ""
            .G_Paterno = row("G_PATERNO") & ""
            .G_Materno = row("G_MATERNO") & ""
            .G_Sexo = row("G_SEXO") & ""
            .G_Calle = row("G_CALLE") & ""
            .G_Nro = row("G_Nro") & ""
            .G_Referencia = row("G_Referencia") & ""
            .G_Comuna = row("G_COMUNA") & ""
            .G_Ciudad = row("G_CIUDAD") & ""
            .G_Email = row("G_EMAIL") & ""
            .G_TipoContrato = row("G_TIPOCONTRATO") & ""
            .G_Plan = row("G_PLAN") & ""
            .G_Obs_Agen = row("G_OBS_AGEN") & ""
            .G_Dato1 = row("G_DATO1") & ""
            .G_MEDIO_PAGO = row("G_MEDIO_PAGO") & ""
            .G_BANCO = row("G_BANCO") & ""
            .G_VENCIMIENTO_TARJETA = row("G_VENCIMIENTO_TARJETA") & ""
            .G_NUMERO_TARJETA = row("G_NUMERO_TARJETA") & ""
            .G_TIPO_TARJETA = row("G_TIPO_TARJETA") & ""
            .G_DIA_PAGO = row("G_DIA_PAGO") & ""

        End With

        inicializarGestionRecuperacion = GESTION

    End Function

    Public Function inicializarClienteRecuperacion(ByRef row As DataRow) As eCliente

        With CLIENTE
            '----------------------------
            .C_Id = row("C_ID") & ""
            .C_Fono1 = row("C_FONO1") & ""
            .C_Telefono1 = row("C_TELEFONO1") & ""
            .C_Fono2 = row("C_FONO2") & ""
            .C_Telefono2 = row("C_TELEFONO2") & ""
            .C_Fono3 = row("C_FONO3") & ""
            .C_Telefono3 = row("C_TELEFONO3") & ""
            .C_Fono4 = row("C_FONO4") & ""
            .C_Telefono4 = row("C_TELEFONO4") & ""
            .C_Fono5 = row("C_FONO5") & ""
            .C_Telefono5 = row("C_TELEFONO5") & ""
            .C_Fono6 = row("C_FONO6") & ""
            .C_Telefono6 = row("C_TELEFONO6") & ""
            .C_Fono_Alternativo = row("C_FONOALTERNATIVO") & ""
            .C_Intentos_Max = row("C_INTENTOS_MAX")
            .C_TipoBase = row("C_TIPOBASE") & ""

            .Extra1 = row("C_extra1") & ""
            .Extra2 = row("C_extra2") & ""
            .Extra3 = row("C_extra3") & ""
            .Extra4 = row("C_extra4") & ""
            .Extra5 = row("C_extra5") & ""
            .Extra6 = row("C_extra6") & ""
            .Extra7 = row("C_extra7") & ""
            .Extra8 = row("C_extra8") & ""
            .Extra9 = row("C_extra9") & ""
            .Extra10 = row("C_extra10") & ""
            .Extra11 = row("C_extra11") & ""
            .Extra12 = row("C_extra12") & ""
            .Extra13 = row("C_extra13") & ""
            .Extra14 = row("C_extra14") & ""
            .Extra15 = row("C_extra15") & ""
            .Extra16 = row("C_extra16") & ""
            .Extra17 = row("C_extra17") & ""
            .Extra18 = row("C_extra18") & ""
            .Extra19 = row("C_extra19") & ""
            .Extra20 = row("C_extra20") & ""
            .Extra21 = row("C_extra21") & ""
            .Extra22 = row("C_extra22") & ""
            .Extra23 = row("C_extra23") & ""
            .Extra24 = row("C_extra24") & ""
            .Extra25 = row("C_extra25") & ""
            .Extra26 = row("C_extra26") & ""
            .Extra27 = row("C_extra27") & ""
            .Extra28 = row("C_extra28") & ""
            .Extra29 = row("C_extra29") & ""
            .Extra30 = row("C_extra30") & ""
            .Extra31 = row("C_extra31") & ""
            .Extra32 = row("C_extra32") & ""
            .Extra33 = row("C_extra33") & ""
            .Extra34 = row("C_extra34") & ""
            .Extra35 = row("C_extra35") & ""
            .Extra36 = row("C_extra36") & ""
            .Extra37 = row("C_extra37") & ""
            .Extra38 = row("C_extra38") & ""
            .Extra39 = row("C_extra39") & ""
            .Extra40 = row("C_extra40") & ""

        End With

        inicializarClienteRecuperacion = CLIENTE


        If IsDBNull(row("G_INTENTO")) Then
            RECUPERACION.R_Intento = 1
        Else
            RECUPERACION.R_Intento = row("G_Intento") + 1
        End If

    End Function


End Module

