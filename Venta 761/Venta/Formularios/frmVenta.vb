Option Strict Off
Option Explicit On
Imports System.Collections.Generic
Imports Entidad
Imports BI

Friend Class frmVenta
    Inherits System.Windows.Forms.Form
    Public KeyAscii As Short

    Public FilaAgregar2 As Integer
    Public FilaElimina As Integer
    Public FilaElimina2 As Integer
    Public fila As Integer
    Public fila2 As Integer
    Public Col As Integer
    Public i As Integer

    Public TpoContratoAdicional As eTipoContrato
    Dim biClsComuna As New clsComunaBI
    Dim biClsCiudad As New clsCiudadBI
    Dim biClsEdoFono As New clsEstadoFonosBI
    Dim biClsScript As New clsScriptBI
    Dim biClsTipoContrato As New clsTipoContratoBI
    Dim biClsPlan As New clsPlanBI
    Dim biClsRestricion As New clsRestriccionBI
    Dim biClsParentesco As New clsParentescoBI
    Dim biClsParentescoCampania As New clsParentescoCampaniaBI
    Dim biClsBen As New clsBeneficiarioBI
    Dim biClsAdic As New clsAdicionalBI
    Dim biCorreoInv As New clsCorreoInvalidoBI

    Dim clsScript As New eScript
    Dim vlUF As String

    Dim ListTipoContrato As New List(Of eTipoContrato)
    Dim medioPagoGlobal As New eMedioPago

    Dim listPlanes As New List(Of ePlan)
    'Dim planE As New ePlan
    'Dim ePlanGlobal As New ePlan
    Dim listRestricciones As New List(Of eRestriccion)
    Dim restricionE As New eRestriccion
    Dim listParentesco As List(Of eParentesco)
    Dim listaCorreoInvalido As New List(Of eCorreoInvalido)

    Private IsInitializing As Boolean = True

    Dim biCliente As New clsClienteBI
    Dim biGeneral As New clsGeneralBI
    Dim biScrisp As New clsScriptBI
    Dim biGesRes As New clsRegrabacionesBI

    Private Sub Botones(ByRef activo As Boolean)
        Select Case activo
            Case False
                CmdLlamar1.Enabled = False
                CmdLlamar2.Enabled = False
                CmdLlamar3.Enabled = False
                CmdLlamar4.Enabled = False
                CmdLlamar5.Enabled = False
                CmdLlamar6.Enabled = False
                CmdLlamarAlt.Enabled = False
            Case True
                CmdLlamar1.Enabled = Not esVacio((Txt_Fono1.Text)) 'True
                CmdLlamar2.Enabled = Not esVacio((Txt_Fono2.Text)) 'True
                CmdLlamar3.Enabled = Not esVacio((Txt_Fono3.Text)) 'True
                CmdLlamar4.Enabled = Not esVacio((Txt_Fono4.Text)) 'True
                CmdLlamar5.Enabled = Not esVacio((Txt_Fono5.Text)) 'True
                CmdLlamar6.Enabled = Not esVacio((Txt_Fono6.Text)) 'True
                CmdLlamarAlt.Enabled = Not esVacio((Txt_Fono_alt.Text)) 'True
                CmdLlamarAlt.Enabled = IIf(perfil = "Regrabador", True, Not esVacio((Txt_Fono_alt.Text)))  'True
                Txt_Fono_alt.ReadOnly = IIf(perfil = "Regrabador", False, True)
                CmdLlamarVent.Enabled = IIf(perfil = "Regrabador", True, Not esVacio((txt_FonoVenta.Text)))  'True
                txt_FonoVenta.ReadOnly = IIf(perfil = "Regrabador", False, True)

        End Select
    End Sub

    'UPGRADE_WARNING: El evento chkMute.CheckStateChanged se puede desencadenar cuando se inicializa el formulario. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub chkMute_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkMute.CheckStateChanged
        If chkMute.CheckState = 0 Then
            chkMute.Text = "Mute Desactivado"
        Else
            chkMute.Text = "Mute Activado"
            chkMute.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
        End If

        If CmdLlamar1.Text = "COLGAR" Or CmdLlamar2.Text = "COLGAR" Or CmdLlamar3.Text = "COLGAR" Or CmdLlamar4.Text = "COLGAR" Or CmdLlamar5.Text = "COLGAR" Or CmdLlamar6.Text = "COLGAR" Or CmdLlamarAlt.Text = "COLGAR" Then
            Datos = ""
            Mute()
        ElseIf chkMute.CheckState = 1 Then
            MsgBox("Debe llamar antes de pasar al estado MUTE")
            chkMute.CheckState = System.Windows.Forms.CheckState.Unchecked
        End If
    End Sub

    Private Sub llamarProgresivoFijo()


        If flg_progresivo_activado Then

            If (flg_fonoVent) And Not flg_EsCeluVent Then
                'Comenzar (1)
                LlamarFono(CmdLlamarVent, txt_FonoVenta, flg_fonoVent)

            Else

                If (flg_fono1) And Not flg_EsCelu1 Then
                    'Comenzar (1)
                    LlamarFono(CmdLlamar1, Txt_Fono1, flg_fono1)

                Else
                    If (flg_fono2) And Not flg_EsCelu2 Then
                        'Comenzar (2)
                        LlamarFono(CmdLlamar2, Txt_Fono2, flg_fono2)

                    Else
                        If (flg_fono3) And Not flg_EsCelu3 Then
                            'Comenzar (3)
                            LlamarFono(CmdLlamar3, Txt_Fono3, flg_fono3)

                        Else
                            If (flg_fono4) And Not flg_EsCelu4 Then
                                'Comenzar (4)
                                LlamarFono(CmdLlamar4, Txt_Fono4, flg_fono4)

                            Else
                                If (flg_fono5) And Not flg_EsCelu5 Then
                                    'Comenzar (5)
                                    LlamarFono(CmdLlamar5, Txt_Fono5, flg_fono5)
                                Else
                                    If (flg_fono6) And Not flg_EsCelu6 Then
                                        'Comenzar (6)
                                        LlamarFono(CmdLlamar6, Txt_Fono6, flg_fono6)
                                    Else
                                        If (flg_fonoalt) And Not flg_EsCeluAlt Then
                                            LlamarFono(CmdLlamarAlt, Txt_Fono_alt, flg_fonoalt)
                                            If flg_progresivo_activado Then
                                                flg_progresivo_activado = False
                                            End If
                                        Else
                                            llamarProgresivoCelular()
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Else
            ' el progesivo se ha desactivado debido a que el ejecutivo
            ' ha marcado una llamada como contactada o la modalidad de la campaña no es progresiva
        End If

    End Sub

    Private Sub llamarProgresivoCelular()

        If flg_progresivo_activado Then
            If (flg_fonoVent) Then
                'Comenzar (1)
                LlamarFono(CmdLlamarVent, txt_FonoVenta, flg_fonoVent)
            Else

                If (flg_fono1) Then
                    'Comenzar (1)
                    LlamarFono(CmdLlamar1, Txt_Fono1, flg_fono1)
                Else
                    If (flg_fono2) Then
                        'Comenzar (2)
                        LlamarFono(CmdLlamar2, Txt_Fono2, flg_fono2)
                    Else
                        If (flg_fono3) Then
                            'Comenzar (3)
                            LlamarFono(CmdLlamar3, Txt_Fono3, flg_fono3)
                        Else
                            If (flg_fono4) Then
                                'Comenzar (4)
                                LlamarFono(CmdLlamar4, Txt_Fono4, flg_fono4)
                            Else
                                If (flg_fono5) Then
                                    'Comenzar (5)
                                    LlamarFono(CmdLlamar5, Txt_Fono5, flg_fono5)
                                Else
                                    If (flg_fono6) Then
                                        'Comenzar (6)
                                        LlamarFono(CmdLlamar6, Txt_Fono6, flg_fono6)
                                    Else
                                        If (flg_fonoalt) Then
                                            LlamarFono(CmdLlamarAlt, Txt_Fono_alt, flg_fonoalt)
                                            If flg_progresivo_activado Then
                                                flg_progresivo_activado = False
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Else
            ' el progesivo se ha desactivado debido a que el ejecutivo
            ' ha marcado una llamada como contactada o la modalidad de la campaña no es progresiva
        End If

    End Sub

    Private Sub llamarProgresivo()

        If flg_progresivo_activado Then


            If (flg_fono1) Then
                'Comenzar (1)
                LlamarFono(CmdLlamar1, Txt_Fono1, flg_fono1)
            Else
                If (flg_fono2) Then
                    'Comenzar (2)
                    LlamarFono(CmdLlamar2, Txt_Fono2, flg_fono2)
                Else
                    If (flg_fono3) Then
                        'Comenzar (3)
                        LlamarFono(CmdLlamar3, Txt_Fono3, flg_fono3)
                    Else
                        If (flg_fono4) Then
                            'Comenzar (4)
                            LlamarFono(CmdLlamar4, Txt_Fono4, flg_fono4)
                        Else
                            If (flg_fono5) Then
                                'Comenzar (5)
                                LlamarFono(CmdLlamar5, Txt_Fono5, flg_fono5)
                            Else
                                If (flg_fono6) Then
                                    'Comenzar (6)
                                    LlamarFono(CmdLlamar6, Txt_Fono6, flg_fono6)
                                Else
                                    If (flg_fonoalt) Then
                                        LlamarFono(CmdLlamarAlt, Txt_Fono_alt, flg_fonoalt)
                                        If flg_progresivo_activado Then
                                            flg_progresivo_activado = False
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

    End Sub

    'pasamos como parametros el fono a llamar y el valor del telefono
    Public Sub LlamarFono(ByRef CmdLlamar As System.Windows.Forms.Button, ByRef txtFono As System.Windows.Forms.TextBox, ByRef fonoActivo As Boolean)
        Dim cmdFonos(6) As String
        Dim i As Short
        Dim f As Short
        If txtFono.Text <> "" Then
            If CmdLlamar.Text = "COLGAR" Then
                grabarCallId("CORTAR", WS_CALL_ID, (txtFono.Text), claveRegistroActual)
                'cortar llamada fonox
                cortarFonos(CmdLlamar, False)
                'desactivamos el mute
                If chkMute.CheckState = 1 Then chkMute.CheckState = System.Windows.Forms.CheckState.Unchecked
                'volvemos a llamar al fono
                fonoActivo = False
            Else
                Botones(True)
                cmdFonos(0) = CmdLlamar1.Name
                cmdFonos(1) = CmdLlamar2.Name
                cmdFonos(2) = CmdLlamar3.Name
                cmdFonos(3) = CmdLlamar4.Name
                cmdFonos(4) = CmdLlamar5.Name
                cmdFonos(5) = CmdLlamar6.Name
                cmdFonos(6) = CmdLlamarAlt.Name

                f = UBound(cmdFonos)

                For i = 0 To f
                    If cmdFonos(i) <> CmdLlamar.Name Then
                        'validamos que no existan otras llamadas activas
                        If i = 0 Then
                            If (CmdLlamar2.Text = "COLGAR" Or CmdLlamar3.Text = "COLGAR" Or CmdLlamar4.Text = "COLGAR" Or CmdLlamar5.Text = "COLGAR" Or CmdLlamar6.Text = "COLGAR" Or CmdLlamarAlt.Text = "COLGAR") Then
                                MsgBox("No puede realizar otra llamada mientras ya tenga una activa!", MsgBoxStyle.Exclamation, "CORDIALPHONE")
                                Exit Sub
                            End If
                        End If

                        If i = 1 Then
                            If (CmdLlamar1.Text = "COLGAR" Or CmdLlamar3.Text = "COLGAR" Or CmdLlamar4.Text = "COLGAR" Or CmdLlamar5.Text = "COLGAR" Or CmdLlamar6.Text = "COLGAR" Or CmdLlamarAlt.Text = "COLGAR") Then
                                MsgBox("No puede realizar otra llamada mientras ya tenga una activa!", MsgBoxStyle.Exclamation, "CORDIALPHONE")
                                Exit Sub
                            End If
                        End If

                        If i = 2 Then
                            If (CmdLlamar1.Text = "COLGAR" Or CmdLlamar2.Text = "COLGAR" Or CmdLlamar4.Text = "COLGAR" Or CmdLlamar5.Text = "COLGAR" Or CmdLlamar6.Text = "COLGAR" Or CmdLlamarAlt.Text = "COLGAR") Then
                                MsgBox("No puede realizar otra llamada mientras ya tenga una activa!", MsgBoxStyle.Exclamation, "CORDIALPHONE")
                                Exit Sub
                            End If
                        End If

                        If i = 3 Then
                            If (CmdLlamar1.Text = "COLGAR" Or CmdLlamar2.Text = "COLGAR" Or CmdLlamar3.Text = "COLGAR" Or CmdLlamar5.Text = "COLGAR" Or CmdLlamar6.Text = "COLGAR" Or CmdLlamarAlt.Text = "COLGAR") Then
                                MsgBox("No puede realizar otra llamada mientras ya tenga una activa!", MsgBoxStyle.Exclamation, "CORDIALPHONE")
                                Exit Sub
                            End If
                        End If

                        If i = 4 Then
                            If (CmdLlamar1.Text = "COLGAR" Or CmdLlamar2.Text = "COLGAR" Or CmdLlamar3.Text = "COLGAR" Or CmdLlamar4.Text = "COLGAR" Or CmdLlamar6.Text = "COLGAR" Or CmdLlamarAlt.Text = "COLGAR") Then
                                MsgBox("No puede realizar otra llamada mientras ya tenga una activa!", MsgBoxStyle.Exclamation, "CORDIALPHONE")
                                Exit Sub
                            End If
                        End If

                        If i = 5 Then
                            If (CmdLlamar1.Text = "COLGAR" Or CmdLlamar2.Text = "COLGAR" Or CmdLlamar3.Text = "COLGAR" Or CmdLlamar4.Text = "COLGAR" Or CmdLlamar5.Text = "COLGAR" Or CmdLlamarAlt.Text = "COLGAR") Then
                                MsgBox("No puede realizar otra llamada mientras ya tenga una activa!", MsgBoxStyle.Exclamation, "CORDIALPHONE")
                                Exit Sub
                            End If
                        End If

                        If i = 6 Then
                            If (CmdLlamar1.Text = "COLGAR" Or CmdLlamar2.Text = "COLGAR" Or CmdLlamar3.Text = "COLGAR" Or CmdLlamar4.Text = "COLGAR" Or CmdLlamar5.Text = "COLGAR" Or CmdLlamar6.Text = "COLGAR") Then
                                MsgBox("No puede realizar otra llamada mientras ya tenga una activa!", MsgBoxStyle.Exclamation, "CORDIALPHONE")
                                Exit Sub
                            End If
                        End If

                    End If
                Next
                'sino hay otras llamadas activas entonces
                'llamamos al fonox
                If Not llamar((txtFono.Text)) Then
                    MsgBox("No puede realizar otra llamada mientras ya tenga una activa!", MsgBoxStyle.Exclamation, "CORDIALPHONE")
                Else
                    grabarCallId("LLAMAR", WS_CALL_ID, (txtFono.Text), claveRegistroActual)
                    CmdLlamar.Text = "COLGAR"
                    CmdLlamar.BackColor = Color.Red
                    'CmdLlamar.BackColor = System.Drawing.ColorTranslator.FromOle(&HFF)
                    txtCallId.Text = WS_CALL_ID
                End If
            End If
        End If
    End Sub
    '******************Metodo al cambiar item de combobox cmbComunicaCon de tab Conexion****************************************************************
    Private Sub CmbComunicaCon_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbComunicaCon.SelectedIndexChanged
        '0 [No Especificado]
        '1 COMUNICA CON CLIENTE
        '2 COMUNICA CON CLIENTE NO VIGENTE EN METLIFE
        '3 COMUNICA CON CONYUGE
        '4 COMUNICA CON TERCERO VALIDO
        '5 COMUNICA CON REGISTRO NO VALIDO (NO VIVE/NO TRABAJA AHI)

        Select Case CmbComunicaCon.SelectedIndex
            Case 1, 3, 4
                CmbComunicaTercero.SelectedIndex = -1
                Label12.Visible = True
                CmbComunicaTercero.Visible = False
                Label11.Visible = False
                Label13.Visible = False
            Case 2
                CmbComunicaTercero.SelectedIndex = -1
                Label12.Visible = True
                CmbComunicaTercero.Visible = True
                Label11.Visible = False
                Label13.Visible = True

        End Select
    End Sub
    '******************Metodo al cambiar item de combobox cmbConecta de tab Conexion****************************************************************
    Private Sub CmbConecta_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbConecta.SelectedIndexChanged
        If CmbConecta.SelectedIndex = 1 Then
            CmbNoConecta.SelectedIndex = 0
            CmbComunicaCon.SelectedIndex = 0
            CmbComunicaTercero.SelectedIndex = 0
            Label11.Visible = False
            CmbNoConecta.Visible = False
            FrmConex.Visible = True
            Label13.Visible = False
            CmbComunicaTercero.Visible = False

            clsScript = CargaScript(1)
            wbScriptBienvenida.DocumentText = clsScript.contenidoScript

        ElseIf CmbConecta.SelectedIndex = 2 Then
            CmbNoConecta.SelectedIndex = 0
            CmbComunicaCon.SelectedIndex = 0
            CmbComunicaTercero.SelectedIndex = 0
            Label11.Visible = True
            CmbNoConecta.Visible = True
            FrmConex.Visible = False
            Label13.Visible = False
            CmbComunicaTercero.Visible = False

        End If
    End Sub

    Private Sub CmbEstAgenda_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbEstAgenda.SelectedIndexChanged
        If CmbEstAgenda.SelectedIndex = 0 Or CmbEstAgenda.SelectedIndex = 1 Then
            FrmAgendamiento.Visible = True
            CmdSiguienteA.Visible = False
            CmdTerminarA.Visible = True
        Else
            FrmAgendamiento.Visible = False
            CmdSiguienteA.Visible = True
            CmdTerminarA.Visible = False
        End If
    End Sub

    Private Sub cmdAnexos_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        frmAnexos.ShowDialog()
    End Sub
    '************************metodo de boton atras ***************************************************************
    Private Sub CmdAtras_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdAtras.Click

        'If perfil <> "Regrabador" Then
        If MsgBox("Desea Retroceder?", MsgBoxStyle.YesNo, csNombreAplicacion) = MsgBoxResult.Yes Then
            ' reinicializa los controles y las variables que se modificaron
            ' al pasar de una pantalla a otra, al presionar el boton VOLVER
            ' se saca el ultimo elemento ingresado a la pila (pop). Este numero
            ' corresponde tambien al ultimo TAB visitado en el flujo. Sabiendo este dato
            ' podemos resetear lo que se hizo al pasar del tab anterior al actual
            Dim pantallaAnterior As Integer
            Dim pantallaActual As String


            ' saco el ultimo elemento en la pila (ultimo TAB que se visito)
            pantallaAnterior = pilaPop()
            ' guardo el TAB ACTUAL!
            pantallaActual = Cuerpo.TabPages.Item(Cuerpo.SelectedIndex).Name

            Select Case pantallaActual
                Case "_Cuerpo_Conex"
                    CmbConecta.SelectedIndex = 0
                    CmbNoConecta.SelectedIndex = 0
                    CmbComunicaCon.SelectedIndex = 0
                    CmbComunicaTercero.SelectedIndex = 0
                    Label11.Visible = False
                    CmbNoConecta.Visible = False
                    FrmConex.Visible = False
                    Label13.Visible = False
                    CmbComunicaTercero.Visible = False

                Case "_Cuerpo_MtvoLL"

                    lblRealiza.Visible = True
                    CmbRealiza.SelectedIndex = 0
                    CmbRealiza.Visible = True

                Case "_Cuerpo_DatosCli"
                    txtNombreV.Text = Trim(UCase(CLIENTE.C_Nombre))
                    txtPaternoV.Text = Trim(UCase(CLIENTE.C_Paterno))
                    txtMaternoV.Text = Trim(UCase(CLIENTE.C_Materno))
                    txtRutV.Text = Mid(CLIENTE.C_Rut, 1, 4)
                    txtDvV.Text = ""
                    txtEmail.Text = CLIENTE.C_Email
                    txtCalleV.Text = CLIENTE.C_Direccion
                    txtNroV.Text = ""
                    txtReferenciaV.Text = ""
                    cmbComuna.SelectedIndex = -1
                    cmbComuna.Text = ""
                    cmbCiudad.Text = ""
                    txtReferenciaV.Text = ""
                    txtFonoVenta.Text = ""
                    CmPrevision.SelectedIndex = 0
                    'cmbPlan.SelectedIndex = 0
                    cmbCiudad.DataSource = Nothing
                    cmbCiudad.ValueMember = Nothing

                Case "_Cuerpo_MPago"

                Case "_Cuerpo_InfAdic"

                Case "_Cuerpo_Certifica"
                    cmbAceptaPrima.SelectedIndex = 0
                    cmbAceptaContrato.SelectedIndex = 0

                Case "_Cuerpo_UltInfo"



                Case "_Cuerpo_Adicionales"

                Case "_Cuerpo_Objeciones"
                    TxtObj.Text = ""
                    CmbObj.SelectedIndex = 0

                Case "_Cuerpo_Agendar"
                    CmbEstAgenda.SelectedIndex = -1
                    FrmAgendamiento.Visible = False
                    TxtObsA.Text = ""
                    DTAgenFecha2.Value = Today
                    cmbHora.SelectedIndex = -1
                    cmbMin.SelectedIndex = -1
                    CmdTerminarA.Visible = True
                    CmdSiguienteA.Visible = True


                Case "_Cuerpo_Agenda2"
                    TxtObsAgen2.Text = ""
                    DTAgenFecha2.Value = Today
                    CmbHora2.SelectedIndex = -1
                    CmbMin2.SelectedIndex = -1


                Case "_Cuerpo_FinNC"

            End Select

            ' ahora se resetea la pantalla anterior, y tambien
            ' TODOS los campos de usuario que se pudieron haber llenado
            Select Case pantallaAnterior
                Case 0 '_Cuerpo_IngresoCli

                Case 1 ' _Cuerpo_Conex
                    Cuerpo.TabPages.Add(_Cuerpo_Conex)
                    GESTION.G_Contacto = ""
                    GESTION.G_Contacto_Con = ""
                    GESTION.G_No_Contacto = ""



                Case 2 '_Cuerpo_MtvoLL
                    Cuerpo.TabPages.Add(_Cuerpo_MtvoLL)
                    CmbRealiza.SelectedIndex = 0
                    GESTION.G_Contacto_Flujo = ""
                    GESTION.G_Motivo_No_Interesa = ""


                Case 3 '_Cuerpo_DatosCli
                    Cuerpo.TabPages.Add(_Cuerpo_DatosCli)
                    If perfil <> "Regrabador" Then

                        GESTION.G_Nombre = ""
                        GESTION.G_Paterno = ""
                        GESTION.G_Materno = ""
                        GESTION.G_Rut = 0
                        GESTION.G_Dv = ""
                        GESTION.G_Fecha_Nacimiento = ""
                        GESTION.G_Fono_Contacto = ""
                        GESTION.G_Email = ""
                        GESTION.G_Calle = ""
                        GESTION.G_Nro = ""
                        GESTION.G_Referencia = ""
                        GESTION.G_IdComuna = ""
                        GESTION.G_IdCiudad = ""
                        GESTION.G_Comuna = ""
                        GESTION.G_Ciudad = ""
                        GESTION.G_Plan = 0
                        GESTION.G_Prima_Uf = ""
                        GESTION.G_Prima_Pesos = 0
                        GESTION.G_IdComuna = ""
                        GESTION.G_IdComuna = ""
                    End If

                Case 4 '_Cuerpo_Mpago
                    Cuerpo.TabPages.Add(_Cuerpo_MPago)

                Case 5 '_Cuerpo_InfAdicional

                    Cuerpo.TabPages.Add(_Cuerpo_InfAdic)

                Case 6 '10 _Cuerpo_Certifica
                    Cuerpo.TabPages.Add(_Cuerpo_Certifica)
                    cmbAceptaPrima.SelectedIndex = 0
                    cmbAceptaContrato.SelectedIndex = 0


                Case 7 '_Cuerpo_InfLL
                    Cuerpo.TabPages.Add(_Cuerpo_UltInfo)

                Case 9 '_Cuerpo_Adicionales
                    Cuerpo.TabPages.Add(_Cuerpo_Adicionales)

                Case 10 '5 _Cuerpo_Objeciones

                    Cuerpo.TabPages.Add(_Cuerpo_Objeciones)

                Case 11 '6 _Cuerpo_Agendar

                    Cuerpo.TabPages.Add(_Cuerpo_Agendar)

                Case 12 '9 _Cuerpo_Agenda2

                    Cuerpo.TabPages.Add(_Cuerpo_Agenda2)

                Case 13 '8 _Cuerpo_FinNC
                    Cuerpo.TabPages.Add(_Cuerpo_FinNC)

            End Select
            Cuerpo.TabPages.Item(Cuerpo.SelectedIndex).Parent = Nothing
            ' bloquea el boton volver en caso de que este en la primera pantalla
            Me.CmdAtras.Enabled = pantallaAnterior > 1
        End If

    End Sub

    Private Sub CmdLlamar1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdLlamar1.Click
        If CmdLlamar1.Text = "LLAMAR" And Trim$(Txt_Fono1.Text) <> "" Then
            txtCallId.Text = ""
            WS_CALL_ID = ""
        End If
        LlamarFono(CmdLlamar1, Txt_Fono1, flg_fono1)
        Fono_A_Llamar = Txt_Fono1.Text
        'lblNumero.Text = Txt_Fono1.Text
        'lblIdNumero.Text = "1"

    End Sub
    Private Sub CmdLlamar2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdLlamar2.Click
        If CmdLlamar2.Text = "LLAMAR" And Trim$(Txt_Fono2.Text) <> "" Then
            txtCallId.Text = ""
            WS_CALL_ID = ""
        End If
        LlamarFono(CmdLlamar2, Txt_Fono2, flg_fono2)
        Fono_A_Llamar = Txt_Fono2.Text
        'lblNumero.Text = Txt_Fono2.Text
        'lblIdNumero.Text = "2"
    End Sub
    Private Sub CmdLlamar3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdLlamar3.Click
        If CmdLlamar3.Text = "LLAMAR" And Trim$(Txt_Fono3.Text) <> "" Then
            txtCallId.Text = ""
            WS_CALL_ID = ""
        End If
        LlamarFono(CmdLlamar3, Txt_Fono3, flg_fono3)
        Fono_A_Llamar = Txt_Fono3.Text
        'lblNumero.Text = Txt_Fono3.Text
        'lblIdNumero.Text = "3"
    End Sub
    Private Sub CmdLlamar4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdLlamar4.Click
        If CmdLlamar4.Text = "LLAMAR" And Trim$(Txt_Fono4.Text) <> "" Then
            txtCallId.Text = ""
            WS_CALL_ID = ""
        End If
        LlamarFono(CmdLlamar4, Txt_Fono4, flg_fono4)
        Fono_A_Llamar = Txt_Fono4.Text
        'lblNumero.Text = Txt_Fono4.Text
        'lblIdNumero.Text = "4"
    End Sub
    Private Sub CmdLlamar5_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdLlamar5.Click
        If CmdLlamar5.Text = "LLAMAR" And Trim$(Txt_Fono5.Text) <> "" Then
            txtCallId.Text = ""
            WS_CALL_ID = ""
        End If
        LlamarFono(CmdLlamar5, Txt_Fono5, flg_fono5)
        Fono_A_Llamar = Txt_Fono5.Text
        'lblNumero.Text = Txt_Fono5.Text
        'lblIdNumero.Text = "5"
    End Sub
    Private Sub CmdLlamar6_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdLlamar6.Click
        If CmdLlamar6.Text = "LLAMAR" And Trim$(Txt_Fono6.Text) <> "" Then
            txtCallId.Text = ""
            WS_CALL_ID = ""
        End If
        LlamarFono(CmdLlamar6, Txt_Fono6, flg_fono6)
        Fono_A_Llamar = Txt_Fono6.Text
        'lblNumero.Text = Txt_Fono6.Text
        'lblIdNumero.Text = "6"
    End Sub
    Private Sub CmdLlamarAlt_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdLlamarAlt.Click
        If CmdLlamarAlt.Text = "LLAMAR" And Trim$(Txt_Fono_alt.Text) <> "" Then
            txtCallId.Text = ""
            WS_CALL_ID = ""
        End If
        LlamarFono(CmdLlamarAlt, Txt_Fono_alt, flg_fonoalt)
        Fono_A_Llamar = Txt_Fono_alt.Text
        'lblNumero.Text = Txt_Fono_alt.Text
        'lblIdNumero.Text = "alt"
    End Sub
    Private Sub CmdLlamarVent_Click(sender As Object, e As EventArgs) Handles CmdLlamarVent.Click
        If CmdLlamarVent.Text = "LLAMAR" And Trim$(txt_FonoVenta.Text) <> "" Then
            txtCallId.Text = ""
            WS_CALL_ID = ""
        End If
        LlamarFono(CmdLlamarVent, txt_FonoVenta, flg_fonoVent)
        Fono_A_Llamar = txt_FonoVenta.Text
        'lblNumero.Text = txt_FonoVenta.Text
        'lblIdNumero.Text = "vent"
    End Sub


    Public Sub Corta_Anteriores()
        If CmdLlamar1.Text = "COLGAR" Then
            grabarCallId("CORTAR", WS_CALL_ID, (Txt_Fono1.Text), claveRegistroActual)
            cortarFonos(CmdLlamar1)

        End If
        If CmdLlamar2.Text = "COLGAR" Then
            grabarCallId("CORTAR", WS_CALL_ID, (Txt_Fono2.Text), claveRegistroActual)
            cortarFonos(CmdLlamar2)

        End If
        If CmdLlamar3.Text = "COLGAR" Then
            grabarCallId("CORTAR", WS_CALL_ID, (Txt_Fono3.Text), claveRegistroActual)
            cortarFonos(CmdLlamar3)

        End If
        If CmdLlamar4.Text = "COLGAR" Then
            grabarCallId("CORTAR", WS_CALL_ID, (Txt_Fono4.Text), claveRegistroActual)
            cortarFonos(CmdLlamar4)

        End If
        If CmdLlamar5.Text = "COLGAR" Then
            grabarCallId("CORTAR", WS_CALL_ID, (Txt_Fono5.Text), claveRegistroActual)
            cortarFonos(CmdLlamar5)

        End If
        If CmdLlamar6.Text = "COLGAR" Then
            grabarCallId("CORTAR", WS_CALL_ID, (Txt_Fono6.Text), claveRegistroActual)
            cortarFonos(CmdLlamar6)

        End If
        If CmdLlamarAlt.Text = "COLGAR" Then
            grabarCallId("CORTAR", WS_CALL_ID, (Txt_Fono_alt.Text), claveRegistroActual)
            cortarFonos(CmdLlamarAlt)

        End If
        If CmdLlamarVent.Text = "COLGAR" Then
            grabarCallId("CORTAR", WS_CALL_ID, (txt_FonoVenta.Text), claveRegistroActual)
            cortarFonos(CmdLlamarVent)

        End If

    End Sub
    ''' <summary>
    ''' procedimiento para buscar cliente en la tabla cli
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Buscar_Cliente()
        tmrPostview.Enabled = False
        Call Corta_Anteriores()
        Dim Tabla As New Data.DataTable

        Tabla = biCliente.Buscar_cliente(WS_IDUSUARIO) 'SI NO HAY NINGUN REGISTRO A EVALUAR SE TERMINA LA APLICACION
        If Tabla.Rows.Count <= 0 Then
            End
        End If

        For x As Integer = 0 To Tabla.Rows.Count - 1
            claveRegistroActual = Tabla.Rows(x)("C_ID")
        Next

        CLIENTE = inicializarCliente(Tabla.Rows(0))

        flg_progresivo_activado = True

        WS_CALL_ID = ""
        txtCallId.Text = ""

        Txt_Fono1.Text = Trim(CLIENTE.C_Telefono1)
        flg_fono1 = Not esVacio(CLIENTE.C_Telefono1)
        flg_EsCelu1 = esCelular(CLIENTE.C_Telefono1)

        Txt_Fono2.Text = Trim(CLIENTE.C_Telefono2)
        flg_fono2 = Not esVacio(CLIENTE.C_Telefono2)
        flg_EsCelu2 = esCelular(CLIENTE.C_Telefono2)

        Txt_Fono3.Text = Trim(CLIENTE.C_Telefono3)
        flg_fono3 = Not esVacio(CLIENTE.C_Telefono3)
        flg_EsCelu3 = esCelular(CLIENTE.C_Telefono3)

        Txt_Fono4.Text = Trim(CLIENTE.C_Telefono4)
        flg_fono4 = Not esVacio(CLIENTE.C_Telefono4)
        flg_EsCelu4 = esCelular(CLIENTE.C_Telefono4)

        Txt_Fono5.Text = Trim(CLIENTE.C_Telefono5)
        flg_fono5 = Not esVacio(CLIENTE.C_Telefono5)
        flg_EsCelu5 = esCelular(CLIENTE.C_Telefono5)

        Txt_Fono6.Text = Trim(CLIENTE.C_Telefono6)
        flg_fono6 = Not esVacio(CLIENTE.C_Telefono6)
        flg_EsCelu6 = esCelular(CLIENTE.C_Telefono6)

        Txt_Fono_alt.Text = Trim(CLIENTE.C_Fono_Alternativo)
        flg_fonoalt = Not esVacio(CLIENTE.C_Fono_Alternativo)
        flg_EsCeluAlt = esCelular(CLIENTE.C_Fono_Alternativo)

        inicializarControles()
        'llamarProgresivoFijo()
        llamarProgresivo()
        tmrPostview.Enabled = True
        lblHora.Text = "0"
        lblMinutos.Text = "00"
        lblSegundos.Text = "00"
        cargaCamposAdic()

    End Sub

    Private Sub cargaCamposAdic()
        Dim ListaCamposAdic As New List(Of eCampoAdicional)
        Dim BiCampoAdic As New clsCampoAdicionalBI
        Dim style As New DataGridViewCellStyle

        ListaCamposAdic = BiCampoAdic.BuscaDatosCampoAdicional(vgCampania.idCRM, CLIENTE.C_Id)
        dtgCamposAdicionales.DataSource = ListaCamposAdic
        dtgCamposAdicionales.Columns(0).Visible = False
        dtgCamposAdicionales.Columns(3).Visible = False
        dtgCamposAdicionales.Columns(4).Visible = False
        dtgCamposAdicionales.Columns(5).Visible = False
        dtgCamposAdicionales.Columns(6).Visible = False

        dtgCamposAdicionales.ColumnHeadersVisible = False
        dtgCamposAdicionales.RowHeadersVisible = False

        dtgCamposAdicionales.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToDisplayedHeaders
        dtgCamposAdicionales.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        style.Font = New Font(dtgCamposAdicionales.Font, FontStyle.Bold)
        dtgCamposAdicionales.Columns(1).DefaultCellStyle = style
        dtgCamposAdicionales.BorderStyle = BorderStyle.Fixed3D
        dtgCamposAdicionales.BackgroundColor = Color.White

    End Sub

    Public Sub cortarFonos(ByRef cmdBoton As System.Windows.Forms.Button, Optional ByRef ConGestion As Boolean = True)
        Dim respuesta As Short
        If Not colgar(ConGestion) Then MsgBox("Se ha detectado un problema al intentar colgar los fonos activos!", MsgBoxStyle.Exclamation, "CORDIALPHONE")
        Select Case db_central
            Case 1, 3
                If txtCallId.Text <> "" Then
                    If ConGestion = False Then
                        respuesta = MsgBox("¿La llamada realizada fue venta?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Responder")
                        If respuesta = 6 Then
                            Botones(False)
                        End If
                    Else
                        If CDbl(GESTION.G_Venta) = 1 Then
                            respuesta = 6
                        End If
                    End If
                    Tiempo(respuesta)
                    cmdBoton.Text = "LLAMAR"
                    cmdBoton.BackColor = Color.LimeGreen

                End If
            Case 2, 4
                cmdBoton.Text = "LLAMAR"
                cmdBoton.BackColor = Color.LimeGreen

        End Select

    End Sub
    '*******INICIALIZAMOS LOS CONTROLES PARA REGRABACION***********
    ''' <summary>
    ''' INICIALIZAMOS LOS CONTROLES PARA REGRABACION
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub inicializarControlesGes()

        Try

            TxtId.Text = Trim(CLIENTE.C_Id)
            TxtRut.Text = GESTION.G_Rut
            txtNombre.Text = GESTION.G_Nombre & " " & Trim(GESTION.G_Paterno) & " " & Trim(GESTION.G_Materno)

            If GESTION.G_Sexo = "M" Then
                CmbSexo.SelectedIndex = 1
            ElseIf GESTION.G_Sexo = "F" Then
                CmbSexo.SelectedIndex = 2
            Else
                CmbSexo.SelectedIndex = 0
            End If

            txtDireccion.Text = Trim(GESTION.G_Calle) & " " & Trim(GESTION.G_Nro) & " " & Trim(GESTION.G_Referencia)
            txtFechaNacimiento.Text = Format(Trim(GESTION.G_Fecha_Nacimiento), "Short Date")

            txtIntentos.Text = CStr(CShort(Val(RECUPERACION.R_Intento)))
            If (Len(Trim(IIf(IsDBNull(GESTION.G_Obs_Agen) = True, "", Trim(GESTION.G_Obs_Agen)))) > 0) Then
                txtObsAgen.Text = Trim(GESTION.G_Obs_Agen)
            End If

            If perfil = "Regrabador" Then
                Txt_Fono_alt.ReadOnly = False
            Else
                Txt_Fono_alt.ReadOnly = True
            End If

            '****inicializamos las variables con la fecha y hora actuales*******
            RECUPERACION.R_Fecha = Today.ToString("yyyyMMdd")
            RECUPERACION.R_Hora = TimeOfDay.ToString("HHmmss")
            RECUPERACION.R_Ejecutivo = WS_IDUSUARIO
            inicializar_controles_tab()


        Catch ex As Exception
            MsgBox("Error: " & ex.Message, MsgBoxStyle.Exclamation, csNombreAplicacion)
        End Try

        tmLabelRegrabacion.Enabled = True
        lblRegrabacion.Visible = True

        cargaCamposAdic()

    End Sub

    '*******INICIALIZAMOS LOS CONTROLES VENTA***********
    ''' <summary>
    ''' INICIALIZAMOS LOS CONTROLES VENTA
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub inicializarControles()

        txtObsAgen.Text = "" : txtIntentos.Text = ""
        txtNombre.Text = ""
        TxtId.Text = Trim(CLIENTE.C_Id)
        'TxtRut.Text = Mid(CLIENTE.C_Rut, 1, 4)
        TxtRut.Text = CLIENTE.C_Rut
        txtNombre.Text = CLIENTE.C_Nombre

        txtNombre.Text = CLIENTE.C_Nombre & " " & Trim(CLIENTE.C_Paterno) & " " & Trim(CLIENTE.C_Materno)

        If CLIENTE.C_Sexo = "M" Then
            CmbSexo.SelectedIndex = 1
        ElseIf CLIENTE.C_Sexo = "F" Then
            CmbSexo.SelectedIndex = 2
        Else
            CmbSexo.SelectedIndex = 0
        End If

        valorPesosUf = CInt(biGeneral.Buscar_Uf)

        txtDireccion.Text = Trim(CLIENTE.C_Direccion)
        txtFechaNacimiento.Text = Format(Trim(CLIENTE.C_Fecha_Nacimiento), "Short Date")

        txtIntentos.Text = CStr(CShort(Val(GESTION.G_Intento)))
        'UPGRADE_WARNING: Se detectó el uso de Null/IsNull(). Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        If (Len(Trim(IIf(IsDBNull(GESTION.G_Obs_Agen) = True, "", Trim(GESTION.G_Obs_Agen)))) > 0) Then
            txtObsAgen.Text = Trim(GESTION.G_Obs_Agen)
        End If

        GESTION.G_Fecha = Today.ToString("yyyyMMdd")
        GESTION.G_Hora = TimeOfDay.ToString("HHmmss")

        GESTION.G_Ejecutivo = Trim(WS_IDUSUARIO)
        GESTION.G_Ip_Ejecutivo = Trim(usuario_actual.IP)
        Fono_A_Llamar = ""

        If Trim(CLIENTE.C_Fecha_Nacimiento) <> "" Then
            txtFechaNacimiento.Text = Mid(CLIENTE.C_Fecha_Nacimiento, 7, 2) & "-" & Mid(CLIENTE.C_Fecha_Nacimiento, 5, 2) & "-" & Mid(CLIENTE.C_Fecha_Nacimiento, 1, 4)
            CLIENTE.C_Edad = CStr(edad(CDate(Mid(CLIENTE.C_Fecha_Nacimiento, 7, 2) & "-" & Mid(CLIENTE.C_Fecha_Nacimiento, 5, 2) & "-" & Mid(CLIENTE.C_Fecha_Nacimiento, 1, 4))))
            dtFechaNacV.Value = CDate(Mid(CLIENTE.C_Fecha_Nacimiento, 7, 2) & "-" & Mid(CLIENTE.C_Fecha_Nacimiento, 5, 2) & "-" & Mid(CLIENTE.C_Fecha_Nacimiento, 1, 4))
        Else
            dtFechaNacV.Value = DateAdd(DateInterval.Year, -30, DateAdd(DateInterval.Day, 1, Now))
        End If
        'TxtEdad.Text = Trim(CLIENTE.c_Edad)
        txtObsAgen.Text = Trim(GESTION.G_Obs_Agen)

        If perfil = "Regrabador" Then
            Txt_Fono_alt.ReadOnly = False
        Else
            Txt_Fono_alt.ReadOnly = True
        End If


        GESTION.G_Contacto = ""
        GESTION.G_No_Contacto = ""
        GESTION.G_Contacto_Con = ""
        GESTION.G_Contacto_Flujo = ""
        GESTION.G_Motivo_No_Interesa = ""
        GESTION.G_Rut = ""
        GESTION.G_Dv = ""
        GESTION.G_Nombre = ""
        GESTION.G_Paterno = ""
        GESTION.G_Materno = ""
        GESTION.G_Fecha_Nacimiento = ""
        GESTION.G_Calle = ""
        GESTION.G_Nro = ""
        GESTION.G_Referencia = ""
        GESTION.G_Comuna = ""
        GESTION.G_Ciudad = ""
        GESTION.G_Email = ""

        GESTION.G_Fecha_Vta = ""
        GESTION.G_Hora_Vta = ""
        GESTION.G_Fecha_Agen = ""
        GESTION.G_Hora_Agen = ""
        GESTION.G_Obs_Agen = ""

        inicializar_controles_tab()
    End Sub
    ' inicializar controles de gestion en todos los tabs utilizados
    ''' <summary>
    ''' inicializar controles de gestion en todos los tabs utilizados
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub inicializar_controles_tab()

        ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

        'TAB 1
        CmbConecta.Focus()
        CmbConecta.SelectedIndex = 0
        CmbNoConecta.SelectedIndex = 0
        CmbComunicaCon.SelectedIndex = 0
        CmbComunicaTercero.SelectedIndex = 0
        Label11.Visible = False
        CmbNoConecta.Visible = False
        FrmConex.Visible = False
        Label13.Visible = False
        CmbComunicaTercero.Visible = False


        'TAB 2

        lblRealiza.Visible = True
        CmbRealiza.SelectedIndex = 0
        CmbRealiza.Visible = True

        'TAB 3

        txtNombreV.Text = ""
        txtPaternoV.Text = ""
        txtMaternoV.Text = ""
        txtRutV.Text = ""
        txtDvV.Text = ""
        txtCelular.Text = ""
        txtEmail.Text = ""
        txtCalleV.Text = ""
        txtNroV.Text = ""
        txtReferenciaV.Text = ""
        txtFonoVenta.Text = ""

        'cmbTipoContrato.SelectedIndex = 0
        If perfil <> "Regrabador" Then
            cmbPlan.SelectedIndex = -1
            If dtAdicional.Rows.Count > 0 Then
                dtAdicional.Rows.Clear()
            End If

        End If
        CmPrevision.SelectedIndex = 0
        cmbComuna.SelectedIndex = -1
        cmbComuna.SelectedText = ""
        cmbCiudad.SelectedIndex = -1
        cmbCiudad.SelectedText = ""
        cmbCiudad.SelectedText = ""
        cmbCiudad.DataSource = Nothing
        cmbCiudad.ValueMember = Nothing

        'TAB 4
        txtNombreA.Text = ""
        txtPaternoA.Text = ""
        txtMaternoA.Text = ""
        txtRutA.Text = ""
        txtDvA.Text = ""
        cmbParentescoAdic.SelectedIndex = -1
        dtFechaNacAdic.Value = dtFechaNacV.MinDate
        dtFechaNacAdic.MaxDate = DateAdd(DateInterval.Day, -6, Now)




        'TAB 5
        CmbObj.SelectedIndex = 0
        FrmObj.Visible = False
        TxtObj.Text = ""
        If perfil = "Regrabador" Then
            Label26.Visible = False
            cmbNoIntereso.Visible = False
        End If


        'TAB 6
        CmbEstAgenda.SelectedIndex = -1
        FrmAgendamiento.Visible = False
        TxtObsA.Text = ""
        DTFechaAgen.Value = Today
        cmbHora.SelectedIndex = -1
        cmbMin.SelectedIndex = -1
        CmdTerminarA.Visible = True
        CmdSiguienteA.Visible = True



        'TAB 9
        TxtObsAgen2.Text = ""
        DTAgenFecha2.Value = Today
        CmbHora2.SelectedIndex = -1
        CmbMin2.SelectedIndex = -1

        'TAB 10
        cmbAceptaPrima.SelectedIndex = -1
        cmbAceptaContrato.SelectedIndex = -1
        lblCargoTarjeta.Visible = True
        lblAcepta.Visible = True
        cmbAceptaPrima.Visible = True
        cmbAceptaContrato.Visible = True
        Panelotro.Visible = False
        txtNumeroCta.Text = ""
        cmbMedioPago.SelectedIndex = -1
        cmbMes.SelectedIndex = -1
        cmbAnio.SelectedIndex = -1


        ' inicializar pila para guardar pantallas visitadas
        ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        pilaInicializar()
        CmdAtras.Enabled = False
        ' (ocultar todos los tabs menos el del inicio)
        Dim i As Integer
        i = Cuerpo.TabCount - 1
        For i = i To 1 Step -1
            Cuerpo.TabPages.Item(i).Parent = Nothing
        Next i
        If Cuerpo.TabPages.Item(0).Name <> "_Cuerpo_Conex" Then
            Cuerpo.TabPages.Item(0).Parent = Nothing
            Cuerpo.TabPages.Add(_Cuerpo_Conex)
        End If


        Cuerpo.Visible = True

        tmLabelRegrabacion.Enabled = False
        lblRegrabacion.Visible = False
    End Sub
    '***********metodo de boton salir de la aplicacion****************
    Private Sub CmdSalir_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdSalir.Click
        Try
            If MsgBox("¿Está seguro que desea salir de la aplicación?", MsgBoxStyle.YesNo, "CORDIALPHONE") = MsgBoxResult.Yes Then
                grabarCallId("CORTAR", WS_CALL_ID, Fono_A_Llamar, claveRegistroActual)


                If Cuerpo.TabPages.Item(0).Name <> "_Cuerpo_IngresoCli" Then
                    If perfil <> "Regrabador" Then
                        Dim x As UInteger
                        x = Convert.ToUInt32(claveRegistroActual)
                        biGeneral.respladar_estado(strQueryUpdateBackupRs, x)
                    End If

                End If
                Logear_Usuario(WS_IDUSUARIO, 2)
                If db_central = 4 Then
                    vpPosicion.LogoutTelefonia((vpPosicion.Usuario))
                End If
                End
            End If

        Catch ex As Exception
            MsgBox(Err.Description & " " & " Error : al salir de la aplicación", MsgBoxStyle.Critical, Me.Text)
            Err.Clear()
        End Try

    End Sub
    ''' <summary>
    ''' Metodo para registrar una venta al cual al cliente no se ha contactado y se da por terminada la gestion
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Terminar()

        GESTION.G_Estado = "T"
        GESTION.G_Venta = CStr(0)
        GESTION.G_Tiempo_PostView = lblHora.Text + ":" + lblMinutos.Text + ":" + lblSegundos.Text

        If perfil <> "Regrabador" Then

            GESTION.G_IdLlamada = WS_CALL_ID
            biCliente.GuardaDatosCliente(CLIENTE, GESTION)
            biCliente.GuardaDatosLog(claveRegistroActual)
            MsgBox("Fin de la gestión. Presione ACEPTAR para continuar con el siguiente registro.", MsgBoxStyle.Information, "CORDIALPHONE")
            limpiarPrimeraPantalla()
            Buscar_Cliente()
        Else
            RECUPERACION.R_Venta = 0
            RECUPERACION.R_Estado = "T"
            RECUPERACION.R_Fecha_Vta = ""
            RECUPERACION.R_Hora_Vta = ""
            RECUPERACION.R_IdLlamada = WS_CALL_ID
            biGesRes.GuardaClienteGes(CLIENTE, RECUPERACION, GESTION)

            MsgBox("Fin de la gestión. Presione ACEPTAR para salir del formulario.", MsgBoxStyle.Information, csNombreAplicacion)
            limpiarPrimeraPantalla()
            Me.Hide()
            frmRegrabaciones.ShowDialog()
            BuscaGes()

        End If
    End Sub
    '******************Metodo al presionar boton Siguiente de tab Conexion****************************************************************
    Private Sub CmdSiguiente_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdSiguiente.Click

        If Trim$(txtCallId.Text) = "" Then
            vpPosicion.CargarPosicion(vpPosicion.Usuario)
            txtCallId.Text = vpPosicion.IDLLAMADA
            WS_CALL_ID = vpPosicion.IDLLAMADA
        End If

        Select Case CmbConecta.SelectedIndex
            Case -1, 0
                MsgBox("Debe selecionar opción si conecta llamada.", vbInformation, "CORDIALPHONE.")
                CmbConecta.Focus()
                Exit Sub

            Case 1

                Select Case CmbComunicaCon.SelectedIndex
                    '0 [No Especificado]
                    '1 COMUNICA CON CLIENTE
                    '2 COMUNICA CON TERCERO VALIDO
                    '3 COMUNICA CON REGISTRO NO VALIDO (NO VIVE/NO TRABAJA AHI)

                    Case -1, 0
                        MsgBox("debe seleccionar con quien se comunica.", vbExclamation, "CORDIALPHONE.")
                        CmbComunicaCon.Focus()
                        Exit Sub

                    Case 1
                        GESTION.G_Contacto = Trim$(CmbConecta.Text)
                        GESTION.G_Contacto_Con = Trim$(CmbComunicaCon.Text)

                        If perfil = "Regrabador" Then
                            cmbReconoce.Visible = True
                            Label36.Visible = True
                            CmbRealiza.Visible = False
                            lblRealiza.Visible = False
                        Else
                            cmbReconoce.Visible = False
                            Label36.Visible = False
                            CmbRealiza.Visible = True
                            lblRealiza.Visible = True
                        End If

                        clsScript = CargaScript(2)
                        wbScriptPresentacion.DocumentText = clsScript.contenidoScript

                        Cuerpo.TabPages.Add(_Cuerpo_MtvoLL)
                        'Cuerpo.TabPages.Add(_Cuerpo_IngresoCli)
                        Cuerpo.TabPages.Item(0).Parent = Nothing
                        guardarPantallaAnterior(1)

                    'Case 3
                    '    GESTION.G_Contacto = Trim$(CmbConecta.Text)
                    '    GESTION.G_Contacto_Con = Trim$(CmbComunicaCon.Text)

                    '    If perfil = "Regrabador" Then
                    '        'CLIENTE.CLI_ESTADO_OBJECION_CALIDAD = 5 'no contactado para regrabacion                            
                    '    End If

                    '    LblFinNoC.Text = ScriptLblFinNoC()
                    '    Cuerpo.TabPages.Add(_Cuerpo_FinNC)
                    '    Cuerpo.TabPages.Item(0).Parent = Nothing
                    '    guardarPantallaAnterior(1)

                    'Case 4

                    '    GESTION.G_Contacto = Trim$(CmbConecta.Text)
                    '    GESTION.G_Contacto_Con = Trim$(CmbComunicaCon.Text)

                    '    If perfil = "Regrabador" Then
                    '        Cuerpo.TabPages.Add(_Cuerpo_Agendar)
                    '    Else
                    '        Cuerpo.TabPages.Add(_Cuerpo_Agenda2)
                    '    End If

                    '    Cuerpo.TabPages.Item(0).Parent = Nothing
                    '    guardarPantallaAnterior(1)

                    Case 2

                        Select Case CmbComunicaTercero.SelectedIndex
                            '0  [No Especificado]
                            '1  TERCERO PIDE DEJAR PENDIENTE
                            '2  VIAJE
                            '3  FALLECIDO
                            '4  NO VIVE AHÍ
                            '5  PROBLEMA POR HORARIO
                            '6  NO DESEA CONTESTAR

                            '0  [No Especificado]
                            '1 De Vacaciones
                            '2 Fallecido
                            '3 Teléfono no corresponde
                            '4 Tercero pide dejar pendiente
                            '5 Viaje
                            '6 Vive fuera
                            '7 Ya no trabaja ahí

                            Case -1, 0
                                MsgBox("Debe seleecionar motivo No comunica.", vbExclamation, "CORDIALPHONE.")
                                CmbComunicaTercero.Focus()
                                Exit Sub

                            Case 1, 4, 5
                                GESTION.G_Contacto = Trim$(CmbConecta.Text)
                                GESTION.G_Contacto_Con = Trim$(CmbComunicaCon.Text)
                                GESTION.G_Contacto_Flujo = Trim$(CmbComunicaTercero.Text)
                                Cuerpo.TabPages.Add(_Cuerpo_Agendar)
                                Cuerpo.TabPages.Item(0).Parent = Nothing
                                guardarPantallaAnterior(1)

                            Case 2, 3, 6, 7
                                If perfil = "Regrabador" Then
                                    '   CLIENTE.CLI_ESTADO_OBJECION_CALIDAD = 5 'no contactado para regrabacion                                    
                                End If
                                GESTION.G_Contacto = Trim$(CmbConecta.Text)
                                GESTION.G_Contacto_Con = Trim$(CmbComunicaCon.Text)
                                GESTION.G_Contacto_Flujo = Trim$(CmbComunicaTercero.Text)

                                LblFinNoC.Text = ScriptLblFinNoC()
                                Cuerpo.TabPages.Add(_Cuerpo_FinNC)
                                Cuerpo.TabPages.Item(0).Parent = Nothing
                                guardarPantallaAnterior(1)
                        End Select
                End Select

            Case 2
                '0[No Especificado]
                '1 OCUPADO
                '2 FUERA DE SERVICIO
                '3 BUZÓN DE VOZ
                '4 NÚMERO NO VÁLIDO
                '5 NO CONTESTA
                '6 FAX O MODEM
                '7 CONGESTIONADO
                '8 FUERA DE ÁEREA O APAGADO


                '1 Buzón de Voz
                '2 Fuera de Area o apagado
                '3 Número No Valido
                '4 No Contesta
                '5 Ocupado

                Select Case CmbNoConecta.SelectedIndex
                    Case -1, 0
                        MsgBox("Debe seleccionar el motivo por el cual No Conecta.", vbExclamation, "CORDIALPHONE.")
                        CmbNoConecta.Focus()
                        Exit Sub

                    Case 99
                        'Case 1, 4, 5
                        GESTION.G_Contacto = Trim$(CmbConecta.Text)
                        GESTION.G_No_Contacto = Trim$(CmbNoConecta.Text)

                        If perfil = "Regrabador" Then
                            Cuerpo.TabPages.Add(_Cuerpo_Agendar)
                        Else
                            Cuerpo.TabPages.Add(_Cuerpo_Agenda2)
                        End If

                        Cuerpo.TabPages.Item(0).Parent = Nothing
                        guardarPantallaAnterior(1)

                    Case Else
                        'Case 2, 3, 6
                        GESTION.G_Contacto = Trim$(CmbConecta.Text)
                        GESTION.G_No_Contacto = Trim$(CmbNoConecta.Text)
                        Terminar()
                End Select
        End Select

    End Sub
    '************************************METODO DE BOTON SIGUIENTE TAB MOTIVO LLAMADO******************************
    Private Sub CmdSiguiente1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdSiguiente1.Click

        If db_central = 4 And Trim(txtCallId.Text) = "" Then
            vpPosicion.CargarPosicion((vpPosicion.Usuario))
            txtCallId.Text = vpPosicion.IDLLAMADA
            WS_CALL_ID = vpPosicion.IDLLAMADA
        End If
        If perfil <> "Regrabador" Then

            '0 [No Especificado]
            '1 Interesa
            '2 No Interesa
            '3 Lo Pensara
            Select Case CmbRealiza.SelectedIndex
                Case 0, -1
                    MsgBox("Selecione opción si cliente interesa Seguro", MsgBoxStyle.Exclamation, csNombreAplicacion)
                    CmbRealiza.Focus()
                    Exit Sub
                Case 1

                    'Si
                    GESTION.G_Contacto_Flujo = Trim(CmbRealiza.Text)
                    Cuerpo.TabPages.Add(_Cuerpo_DatosCli)
                    Cuerpo.TabPages.Item(0).Parent = Nothing
                    llenarTipoContrato()
                    llenar_planes()
                    cmbPlan.Visible = False
                    lblPlanes.Visible = False
                    btnAdicional.Visible = False
                    btnBeneficiarios.Visible = False
                    AsignadatosCli()
                    guardarPantallaAnterior(2)

                Case 2 ' NO
                    GESTION.G_Contacto_Flujo = Trim(CmbRealiza.Text)
                    Cuerpo.TabPages.Add(_Cuerpo_Objeciones)
                    Cuerpo.TabPages.Item(0).Parent = Nothing
                    guardarPantallaAnterior(2)

                Case 3 'lo pensara

                    GESTION.G_Contacto_Flujo = Trim(CmbRealiza.Text)
                    Cuerpo.TabPages.Add(_Cuerpo_Agenda2)
                    Cuerpo.TabPages.Item(0).Parent = Nothing
                    guardarPantallaAnterior(2)
            End Select
        Else

            Select Case cmbReconoce.SelectedIndex
                '0 [No Especificado]
                '1 Si
                '2 No
                Case -1, 0
                    MsgBox("Selecione si cliente reconoce la venta", vbExclamation, "CORDIALPHONE.")
                    cmbReconoce.Focus()
                    Exit Sub
                Case 1
                    RECUPERACION.R_Reconoce = cmbReconoce.Text

                    Select Case CmbRealiza.SelectedIndex
                        Case -1, 0
                            MsgBox("Selecione si cliente acepta seguro", vbExclamation, "CORDIALPHONE.")
                            CmbRealiza.Focus()
                            Exit Sub
                        Case 1
                            GESTION.G_Contacto_Flujo = CmbRealiza.Text
                            Cuerpo.TabPages.Add(_Cuerpo_DatosCli)
                            Cuerpo.TabPages.Item(0).Parent = Nothing
                            llenar_planes()
                            AsignadatosCliGes()
                            guardarPantallaAnterior(2)

                        Case 2 'no interesa
                            '   CLIENTE.CLI_ESTADO_OBJECION_CALIDAD = 3
                            GESTION.G_Contacto_Flujo = CmbRealiza.Text
                            Cuerpo.TabPages.Add(_Cuerpo_Objeciones)
                            Cuerpo.TabPages.Item(0).Parent = Nothing
                            guardarPantallaAnterior(2)

                        Case 3 'Lo pensara
                            GESTION.G_Contacto_Flujo = CmbRealiza.Text
                            Cuerpo.TabPages.Add(_Cuerpo_Agenda2)
                            Cuerpo.TabPages.Item(0).Parent = Nothing
                            guardarPantallaAnterior(2)

                    End Select


                Case 2
                    RECUPERACION.R_Reconoce = cmbReconoce.Text
                    GESTION.G_Contacto_Flujo = ""
                    'CLIENTE.CLI_ESTADO_OBJECION_CALIDAD = 7
                    'CLIENTE.CLI_RECONOCEVTA = ComboBoxReconoce.Text
                    Cuerpo.TabPages.Add(_Cuerpo_Objeciones)
                    Cuerpo.TabPages.Item(0).Parent = Nothing
                    guardarPantallaAnterior(2)

            End Select
        End If
    End Sub

    Private Sub llenarTipoContrato()

        vgListTipoContrato = biClsTipoContrato.ListaTipoContratoPorCampania(vgCampania.idCRM)
        vgFuncionComun.llenaComboBox(cmbTipoContrato, "nombreTipoContrato", "idTipoContrato", vgListTipoContrato.ToArray)

    End Sub

    ''' <summary>
    ''' metodo para hacer visible ciertos controles de el tab Motivo de llamado cuando el perfil sea de regrabador
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ocultar_cmbopcion()
        cmbReconoce.Visible = True
        Label36.Visible = True
    End Sub
    '******************metodo de boton siguiente tab manejo de objeciones ***************************************
    Private Sub CmdSiguiente11_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdSiguiente11.Click
        If perfil <> "Regrabador" Then
            If CmbObj.SelectedIndex <= 0 Then
                MsgBox("Seleccione porque el cliente No desea contratar seguro.", MsgBoxStyle.Exclamation, csNombreAplicacion)
                CmbObj.Focus()
                Exit Sub
            Else
                ' CLIENTE.cli_nointereso = cmbNoIntereso.Text
            End If

            If CmbObj.SelectedIndex <= 0 Then
                MsgBox("Debe ingresar observacion", MsgBoxStyle.Exclamation, csNombreAplicacion)
                CmbObj.Focus()
            Else
                ' CLIENTE.cli_aobsmtvo_nointeresa = TxtObj.Text
                GESTION.G_Motivo_No_Interesa = CmbObj.Text
                LblFinNoC.Text = ScriptLblFinNoC()
                Cuerpo.TabPages.Add(_Cuerpo_FinNC)
                Cuerpo.TabPages.Item(0).Parent = Nothing
                guardarPantallaAnterior(10)
            End If
        Else

            If CmbObj.SelectedIndex <= 0 Then
                MsgBox("Debe ingresar observacion", MsgBoxStyle.Exclamation, csNombreAplicacion)
                CmbObj.Focus()
            Else
                'CLIENTE.CLI_OBSERVACION = TxtObj.Text
                GESTION.G_Motivo_No_Interesa = CmbObj.Text
                LblFinNoC.Text = ScriptLblFinNoC()
                Cuerpo.TabPages.Add(_Cuerpo_FinNC)
                Cuerpo.TabPages.Item(0).Parent = Nothing
                guardarPantallaAnterior(10)
            End If
        End If

    End Sub


    '********************metodo de boton siguiente de tab toma de datos********************************
    Private Sub CmdSiguiente2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdSiguiente2.Click
        If db_central = 4 And Trim(txtCallId.Text) = "" Then
            vpPosicion.CargarPosicion((vpPosicion.Usuario))
            txtCallId.Text = vpPosicion.IDLLAMADA
            WS_CALL_ID = vpPosicion.IDLLAMADA
        End If

        If ValidaInformacion() Then
            GESTION.G_Nombre = Trim(Replace(Trim(UCase(txtNombreV.Text)), "'", "´"))
            GESTION.G_Paterno = Trim(Replace(Trim(UCase(txtPaternoV.Text)), "'", "´"))
            GESTION.G_Materno = Trim(Replace(Trim(UCase(txtMaternoV.Text)), "'", "´"))
            GESTION.G_Rut = Trim(txtRutV.Text)
            GESTION.G_Dv = Trim(txtDvV.Text)
            GESTION.G_Fecha_Nacimiento = dtFechaNacV.Value.ToString("yyyyMMdd")
            GESTION.G_Fono_Contacto = Trim(txtFonoVenta.Text)

            If txtEmail.Text = "" Then
                GESTION.G_Email = ""
            Else
                GESTION.G_Email = Replace(Replace(Trim(txtEmail.Text), "'", "´"), " ", "")
            End If

            GESTION.G_Sexo = CmbSexo.Text
            GESTION.G_Calle = Trim(UCase(Replace(txtCalleV.Text, "'", "`")))
            GESTION.G_Nro = Trim(Replace(txtNroV.Text, "'", "`"))
            GESTION.G_Referencia = Trim$(Replace(txtReferenciaV.Text, "'", "`"))

            'COMUNA CIUDAD REGION
            Dim daCiudad As New clsCiudadBI
            Dim eCiudad As New eCiudad
            GESTION.G_IdComuna = cmbComuna.SelectedValue
            GESTION.G_Comuna = Trim(cmbComuna.Text)
            GESTION.G_IdCiudad = cmbCiudad.SelectedValue
            GESTION.G_Ciudad = Trim(cmbCiudad.Text)

            'PLANES Y TIPO DE CONTRATO|
            Dim ePlan As New ePlan
            ePlan = biClsPlan.BuscarPlanPorIdPlan(cmbPlan.SelectedValue)
            'GESTION.G_Prima_Uf = ePlan.primaUF
            GESTION.G_Prima_Uf = lblPrimaUF.Text
            'GESTION.G_Prima_Pesos = ePlan.PrimaPesos
            GESTION.G_Prima_Pesos = lblPrimaPesos.Text
            GESTION.G_Plan = ePlan.idPlan
            GESTION.G_TipoContrato = ePlan.idTipoContrato
            GESTION.G_Dato1 = CmPrevision.Text

            Dim TpoContratoAdicional As New eTipoContrato
            TpoContratoAdicional = biClsTipoContrato.BuscarTipoContratoPorIdTipoContrato(GESTION.G_TipoContrato)

            If TpoContratoAdicional.cantidadAdicionales <> 0 And lstAdi.Count = 0 Then
                MsgBox("Debe Agregar al menos 1 Adicional", MsgBoxStyle.Information, csNombreAplicacion)
                btnAdicional.Focus()
                Exit Sub
            Else
                If TpoContratoAdicional.definido = True And TpoContratoAdicional.cantidadAdicionales <> lstAdi.Count Then
                    If (TpoContratoAdicional.cantidadAdicionales < lstAdi.Count) Then
                        MsgBox("Debe seleccionar el tipo contrato con la cantidad de  Adicionales, de acuerdo a los adicionales agregados", MsgBoxStyle.Information, csNombreAplicacion)
                        btnAdicional.Focus()
                        Exit Sub
                    End If

                    If (TpoContratoAdicional.definido = True And lstAdi.Count < TpoContratoAdicional.cantidadAdicionales) Then
                        MsgBox("Debe ingresar la cantidad de Adicional(es) de acuerdo al tipo contrato seleccionado", MsgBoxStyle.Information, csNombreAplicacion)
                        btnAdicional.Focus()
                        Exit Sub
                    End If
                End If
            End If

            If TpoContratoAdicional.cantidadBeneficiarios <> lstBen.Count Then
                If (lstBen.Count = 0) Then
                    MsgBox("Debe seleccionar el tipo contrato con la cantidad de " & lstBen.Count & " Beneficiario(s), de acuerdo a los beneficiarios agregados", MsgBoxStyle.Information, csNombreAplicacion)
                    btnBeneficiarios.Focus()
                    Exit Sub
                End If

                'If (lstBen.Count < TpoContratoAdicional.cantidadBeneficiarios) Then
                '    MsgBox("Debe ingresar la cantidad de " & TpoContratoAdicional.cantidadBeneficiarios & " Beneficiario(s) de acuerdo al tipo contrato seleccionado", MsgBoxStyle.Information, csNombreAplicacion)
                '    btnBeneficiarios.Focus()
                '    Exit Sub
                'End If
            End If


            'script información adicional
            clsScript = CargaScript(5)
            wbScriptCertificacion.DocumentText = clsScript.contenidoScript

            Cuerpo.TabPages.Add(_Cuerpo_MPago)

            DatosMedioPago()
            Cuerpo.TabPages.Item(0).Parent = Nothing
            guardarPantallaAnterior(3)
        End If
    End Sub


    Private Sub CmdSiguienteA_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdSiguienteA.Click
        Select Case CmbEstAgenda.SelectedIndex
            Case -1
                MsgBox("Debe seleccionar una opción para el estado del agendamiento.", MsgBoxStyle.Information, csNombreAplicacion)
                CmbEstAgenda.Focus()
                Exit Sub

            Case 2
                If perfil = "Regrabador" Then
                    ' CLIENTE.CLI_ESTADO_OBJECION_CALIDAD = 5 'no contactado para regrabacion
                End If

                'CLIENTE.cli_agen_estado = Trim(CmbEstAgenda.Text)
                LblFinNoC.Text = ScriptLblFinNoC()
                Cuerpo.TabPages.Add(_Cuerpo_FinNC)
                Cuerpo.TabPages.Item(0).Parent = Nothing
                guardarPantallaAnterior(11)
        End Select
    End Sub
    '************METODO DE BOTON TERMINAR DE TAB AGENDAR**********************************
    Private Sub CmdTerminarA_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdTerminarA.Click
        If perfil <> "Regrabador" Then
            Select Case CmbEstAgenda.SelectedIndex
                Case -1
                    MsgBox("Debe seleccionar una opción para el estado del agendamiento.", MsgBoxStyle.Information, csNombreAplicacion)
                    CmbEstAgenda.Focus()
                    Exit Sub
                Case 0, 1
                    'UPGRADE_WARNING: El comportamiento de DateDiff puede ser diferente. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
                    If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(FechaServidor()), DTFechaAgen.Value) > 7 Then
                        MsgBox("La fecha de agendamiento debe ser menor o igual a 1 semana.", MsgBoxStyle.Information, csNombreAplicacion)
                        Exit Sub
                    End If

                    If cmbHora.SelectedIndex = -1 Or cmbMin.SelectedIndex = -1 Then
                        MsgBox("Ingrese hora para agendar nuevo llamado.", MsgBoxStyle.Information, csNombreAplicacion)
                        cmbHora.Focus()
                        Exit Sub
                    Else
                        'CLIENTE.cli_agen_estado = Trim(CmbEstAgenda.Text)
                        GESTION.G_Fecha_Agen = DTFechaAgen.Value.ToString("yyyyMMdd")
                        GESTION.G_Hora_Agen = cmbHora.Text & cmbMin.Text
                        GESTION.G_Obs_Agen = Trim(Replace(TxtObsA.Text, "'", "`"))
                        GESTION.G_IdLlamada = WS_CALL_ID
                        GESTION.G_Venta = 0
                        GESTION.G_Estado = IIf(GESTION.G_Intento >= CLIENTE.C_Intentos_Max, "T", "A")

                        biCliente.GuardaDatosCliente(CLIENTE, GESTION)
                        biCliente.GuardaDatosLog(claveRegistroActual)
                        MsgBox("Fin de la gestión. Presione ACEPTAR para continuar con el siguiente registro.", MsgBoxStyle.Information, csNombreAplicacion)
                        limpiarPrimeraPantalla()
                        Buscar_Cliente()
                    End If


            End Select
        Else
            Select Case CmbEstAgenda.SelectedIndex
                Case -1
                    MsgBox("Debe seleccionar una opción para el estado del agendamiento.", MsgBoxStyle.Information, csNombreAplicacion)
                    CmbEstAgenda.Focus()
                    Exit Sub
                Case 0, 1
                    'UPGRADE_WARNING: El comportamiento de DateDiff puede ser diferente. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
                    If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(FechaServidor()), DTFechaAgen.Value) > 7 Then
                        MsgBox("La fecha de agendamiento debe ser menor o igual a 1 semana.", MsgBoxStyle.Information, csNombreAplicacion)
                        Exit Sub
                    End If

                    If cmbHora.SelectedIndex = -1 Or cmbMin.SelectedIndex = -1 Then
                        MsgBox("Ingrese hora para agendar nuevo llamado.", MsgBoxStyle.Information, csNombreAplicacion)
                        cmbHora.Focus()
                        Exit Sub
                    Else
                        'CLIENTE.cli_agen_estado = Trim(CmbEstAgenda.Text)
                        RECUPERACION.R_Fecha_Agen = DTFechaAgen.Value.ToString("yyyyMMdd")
                        RECUPERACION.R_Hora_Agen = cmbHora.Text & cmbMin.Text
                        RECUPERACION.R_Obs_Agen = Trim(Replace(TxtObsA.Text, "'", "`"))
                        RECUPERACION.R_IdLlamada = WS_CALL_ID
                        RECUPERACION.R_Venta = 0
                        RECUPERACION.R_Estado = IIf(GESTION.G_Intento >= CLIENTE.C_Intentos_Max, "T", "A")
                        ' CLIENTE.CLI_ESTADO_OBJECION_CALIDAD = "2"
                        RECUPERACION.R_Fecha_Vta = ""
                        RECUPERACION.R_Hora_Vta = ""

                        biGesRes.GuardaClienteGes(CLIENTE, RECUPERACION, GESTION)
                        ' biGesRes.ActualizaClienteAgen(CLIENTE.C_Id, CLIENTE.CLI_ESTADO_OBJECION_CALIDAD, CLIENTE.CLI_CALL_ID_CALIDAD)
                        MsgBox("Fin de la gestión. Presione ACEPTAR para salir del formulario.", MsgBoxStyle.Information, csNombreAplicacion)
                        limpiarPrimeraPantalla()
                        Me.Hide()
                        frmRegrabaciones.ShowDialog()
                        BuscaGes()
                    End If
                Case 2
                    'CLIENTE.cli_agen_estado = Trim$(CmbEstAgenda.Text)
                    Terminar()
            End Select
        End If

    End Sub
    'UPGRADE_WARNING: Form evento frmVenta.Activate tiene un nuevo comportamiento. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
    Private Sub frmVenta_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ' verifica si es la primera vez que se activa el formulario
        ' en este caso busca un cliente inmediatamente y luego baja el flag
        If flag_primeravez Then
            flag_primeravez = False
        End If
    End Sub

    Private Sub frmVenta_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        flag_primeravez = True
        ModGeneral.Main()
        ' Me.Text = vgCampania.calNombre & " " & perfil & "NEW Versión: " & My.Application.Info.Version.Major.ToString & "." & My.Application.Info.Version.Minor.ToString _
        '  & "." & My.Application.Info.Version.Revision.ToString

        If perfil = "Regrabador" Then
            frmRegrabaciones.ShowDialog()
        End If

        'asignamos  la fecha maxima ingreso de fecha 
        'dtFechaNac.MaxDate = DateAdd(DateInterval.Year, -18, DateAdd(DateInterval.Day, 1, Now))
        'dtFechaNac.MaxDate = DateAdd(DateInterval.Year, -18, DateAdd(DateInterval.Day, 1, Now))

        'asignamos  la fecha minima ingreso de fecha
        dtFechaNacV.MinDate = dtFechaNacV.MinDate

        Fono_A_Llamar = ""
        vgListEdoFono = biClsEdoFono.listarEstadoFono

        Dim daComuna As New clsComunaBI
        Dim daCiudad As New clsCiudadBI
        vgListComuna = daComuna.listarComuna()
        vgListCiudad = daCiudad.ListaCiudad()
        vgFuncionComun.llenaComboBox(cmbComuna, "nombreComuna", "idComuna", vgListComuna.ToArray)
        cmbComuna.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        cmbComuna.AutoCompleteSource = AutoCompleteSource.ListItems


        Dim listPatertescocontact As New List(Of eTipoParentesco)
        Dim parentesco As New clsParentescoCampaniaBI

        'listParentescoCampania = biClsParentescoCampania.BuscarParentescoPorId(vgCampania.calCodigo, 3)
        'vgFuncionComun.llenaComboBox(cmbParentesco, "nombreParentesco", "idParentesco", listParentescoCampania.ToArray)
        'vgListParentescoCampania = biClsParentescoCampania.BuscarParentescoPorId(vgCampania.calCodigo, 2)

        listParentescoCampania = biClsParentescoCampania.BuscarParentescoPorId(vgCampania.idCRM, 2)

        Dim listMotivoRechazo As New List(Of eMotivoRechazo)
        Dim biMotivoRechazo As New clsMotivoRechazoBI
        listMotivoRechazo = biMotivoRechazo.BuscarMotivoRechazoPorSponsor(vgCampania.idCRM)
        vgFuncionComun.llenaComboBox(CmbObj, "nombreMotivosObjecion", "idMotivosObjecion", listMotivoRechazo.ToArray)

        Dim listExclusion As New List(Of eExclusion)
        Dim biExclusion As New clsExclusionBI
        listExclusion = biExclusion.listarExclusiones(vgCampania.idCRM)
        vgFuncionComun.llenaCheckBox(frmExclusiones.chkListBoxExclusiones, listExclusion)

        listaCorreoInvalido.Clear()
        listaCorreoInvalido = biCorreoInv.listarCorreosInvalido


        If perfil = "Regrabador" Then
            CmbNoConecta.Items.Add("No contesta último intento")
            BuscaGes()
        ElseIf perfil = "Ejecutivo" Then
            Buscar_Cliente()

        End If

        System.Windows.Forms.Application.DoEvents()
    End Sub

    Public Sub BuscaGes()
        Dim dt As New DataTable
        dt = biCliente.Buscar_Gescliente(WS_IDUSUARIO, GesId)
        ''VALIDAR DATO
        flg_progresivo_activado = True

        ' lblEstadoLlamada.Text = ""
        WS_CALL_ID = ""
        txtCallId.Text = ""

        CLIENTE = New eCliente
        GESTION = New eGestion
        RECUPERACION = New eRecuperacion

        CLIENTE = inicializarClienteRecuperacion(dt.Rows(0))
        GESTION = inicializarGestionRecuperacion(dt.Rows(0))

        Rellenar_fonos_Regrabaciones()

        inicializarControlesGes()
        'llamarProgresivoFijo()
        llamarProgresivo()
    End Sub

    Public Sub Rellenar_fonos_Regrabaciones()
        Txt_Fono1.Text = Trim(CLIENTE.C_Telefono1)
        flg_fono1 = Not esVacio(CLIENTE.C_Telefono1)
        flg_EsCelu1 = esCelular(CLIENTE.C_Telefono1)

        Txt_Fono2.Text = Trim(CLIENTE.C_Telefono2)
        flg_fono2 = Not esVacio(CLIENTE.C_Telefono2)
        flg_EsCelu2 = esCelular(CLIENTE.C_Telefono2)

        Txt_Fono3.Text = Trim(CLIENTE.C_Telefono3)
        flg_fono3 = Not esVacio(CLIENTE.C_Telefono3)
        flg_EsCelu3 = esCelular(CLIENTE.C_Telefono3)

        Txt_Fono4.Text = Trim(CLIENTE.C_Telefono4)
        flg_fono4 = Not esVacio(CLIENTE.C_Telefono4)
        flg_EsCelu4 = esCelular(CLIENTE.C_Telefono4)

        Txt_Fono5.Text = Trim(CLIENTE.C_Telefono5)
        flg_fono5 = Not esVacio(CLIENTE.C_Telefono5)
        flg_EsCelu5 = esCelular(CLIENTE.C_Telefono5)

        Txt_Fono6.Text = Trim(CLIENTE.C_Telefono6)
        flg_fono6 = Not esVacio(CLIENTE.C_Telefono6)
        flg_EsCelu6 = esCelular(CLIENTE.C_Telefono6)

        Txt_Fono_alt.Text = Trim(CLIENTE.C_Fono_Alternativo)
        flg_fonoalt = Not esVacio(CLIENTE.C_Fono_Alternativo)
        flg_EsCeluAlt = esCelular(CLIENTE.C_Fono_Alternativo)

        txt_FonoVenta.Text = Trim(GESTION.G_Fono_Contacto)
        flg_fonoVent = Not esVacio(GESTION.G_Fono_Contacto)
        flg_EsCeluVent = esCelular(GESTION.G_Fono_Contacto)

    End Sub

    Private Sub tmCallID_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        If db_central = 4 Then
            If CmdLlamar1.Text = "Colgar" Or CmdLlamar2.Text = "Colgar" Or CmdLlamar3.Text = "Colgar" Or CmdLlamar4.Text = "Colgar" Or CmdLlamar5.Text = "Colgar" Or CmdLlamar6.Text = "Colgar" Or CmdLlamarAlt.Text = "Colgar" Then
                If Trim(txtCallId.Text) = "" Then
                    vpPosicion.CargarPosicion((vpPosicion.Usuario))
                    txtCallId.Text = vpPosicion.IDLLAMADA
                    WS_CALL_ID = vpPosicion.IDLLAMADA

                End If
            End If
        End If
    End Sub

    Private Sub tmrEstadoLlamada_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrEstadoLlamada.Tick
        On Error Resume Next
        'lblEstadoLlamada.Text = EstadoLLamada()
        If db_central = 4 Then
            If Not vpPosicion.EvaluaEstado((vpPosicion.Usuario)) Then
                Corta_Anteriores()
            End If
        End If
    End Sub


    Private Sub TxtObj_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtObj.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        KeyAscii = CaracterValido(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub TxtObsA_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtObsA.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        KeyAscii = CaracterValido(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub TxtObsAgen2_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtObsAgen2.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        KeyAscii = CaracterValido(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub Txt_Fono_alt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Txt_Fono_alt.TextChanged
        Try
            If Txt_Fono_alt.Text <> "" Then
                CLIENTE.C_Fono_Alternativo = Txt_Fono_alt.Text
            Else
                CLIENTE.C_Fono_Alternativo = ""
            End If
        Catch ex As Exception

        End Try

    End Sub
    '*******************METODO AL PRESIONAR BOTON TERMINAR DE TAB FIN NO CONTRATA**********************************
    Private Sub CmdTerminarNC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdTerminarNC.Click
        If perfil <> "Regrabador" Then
            GESTION.G_IdLlamada = WS_CALL_ID
            GESTION.G_Estado = "T"
            GESTION.G_Venta = "0"
            GESTION.G_Tiempo_PostView = lblHora.Text + ":" + lblMinutos.Text + ":" + lblSegundos.Text

            biCliente.GuardaDatosCliente(CLIENTE, GESTION)
            biCliente.GuardaDatosLog(claveRegistroActual)
            MsgBox("Fin de la gestión. Presione ACEPTAR para continuar con el siguiente registro.", MsgBoxStyle.Information, "CORDIALPHONE")
            limpiarPrimeraPantalla()
            Buscar_Cliente()
        Else

            RECUPERACION.R_Venta = 0
            RECUPERACION.R_Estado = "T"
            RECUPERACION.R_Fecha_Vta = ""
            RECUPERACION.R_Hora_Vta = ""
            RECUPERACION.R_IdLlamada = WS_CALL_ID
            RECUPERACION.R_Estado = IIf(RECUPERACION.R_Intento >= CLIENTE.C_Intentos_Max, "T", "A")
            ' CLIENTE.CLI_ESTADO_OBJECION_CALIDAD = "2"

            biGesRes.GuardaClienteGes(CLIENTE, RECUPERACION, GESTION)
            ' biGesRes.ActualizaClienteAgen(CLIENTE.C_Id, CLIENTE.CLI_ESTADO_OBJECION_CALIDAD, CLIENTE.CLI_CALL_ID_CALIDAD)
            MsgBox("Fin de la gestión. Presione ACEPTAR para salir del formulario.", MsgBoxStyle.Information, csNombreAplicacion)
            limpiarPrimeraPantalla()
            Me.Hide()
            frmRegrabaciones.ShowDialog()
            BuscaGes()

        End If

    End Sub

    '*************METODO DE BOTON TERMINAR EN TAB AGENDAR 2********************************
    Private Sub CmdTerminarA2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdTerminarA2.Click
        If CmdLlamar1.Text = "COLGAR" Or CmdLlamar2.Text = "COLGAR" Or CmdLlamar3.Text = "COLGAR" Or CmdLlamar4.Text = "COLGAR" Or CmdLlamar5.Text = "COLGAR" Or CmdLlamar6.Text = "COLGAR" Or CmdLlamarAlt.Text = "COLGAR" Then
            MsgBox("Debe Colgar la llamada antes de terminar", vbExclamation, csNombreAplicacion)
            Exit Sub
        End If

        If perfil <> "Regrabador" Then
            If CmbHora2.SelectedIndex = -1 Or CmbMin2.SelectedIndex = -1 Then
                MsgBox("Debe seleccionar FECHA y HORA ", vbInformation, "CORDIALPHONE")
                CmbHora2.Focus()
                Exit Sub
            Else
                GESTION.G_IdLlamada = WS_CALL_ID
                GESTION.G_Estado = IIf(GESTION.G_Intento >= CLIENTE.C_Intentos_Max, "T", "A")

                If GESTION.G_Contacto_Flujo = "Lo Pensara" Then
                    GESTION.G_Estado = "A"
                End If

                GESTION.G_Obs_Agen = Trim$(Replace(TxtObsAgen2.Text, "'", "´"))
                GESTION.G_Fecha_Agen = DTAgenFecha2.Value.ToString("yyyyMMdd")
                GESTION.G_Hora_Agen = Trim(CmbHora2.Text) & Trim(CmbMin2.Text)
                GESTION.G_Tiempo_PostView = lblHora.Text + ":" + lblMinutos.Text + ":" + lblSegundos.Text

                GESTION.G_Venta = 0
                Dim fecha_agendamiento As String
                fecha_agendamiento = Format(DTAgenFecha2.Value, "dd/MM/yyyy") & " a las " & CmbHora2.Text & ":" & CmbMin2.Text & " Hrs."

                biCliente.GuardaDatosCliente(CLIENTE, GESTION)
                biCliente.GuardaDatosLog(claveRegistroActual)
                MsgBox("Fin de la gestión. Presione ACEPTAR para continuar con el siguiente registro.", vbInformation, "CORDIALPHONE")
                limpiarPrimeraPantalla()
                Buscar_Cliente()
                'Cuerpo.TabPages.Add(_Cuerpo_Conex)
                'Cuerpo.TabPages.Item(0).Parent = Nothing
            End If
        Else
            If CmbHora2.SelectedIndex = -1 Or CmbMin2.SelectedIndex = -1 Then
                MsgBox("Debe seleccionar FECHA y HORA ", vbInformation, "CORDIALPHONE")
                CmbHora2.Focus()
                Exit Sub
            Else


                'CLIENTE.cli_agen_estado = ""
                'CLIENTE.CLI_CALL_ID_CALIDAD = WS_CALL_ID
                RECUPERACION.R_Fecha_Agen = DTFechaAgen.Value.ToString("yyyyMMdd")
                RECUPERACION.R_Hora_Agen = cmbHora.Text & cmbMin.Text
                RECUPERACION.R_Obs_Agen = Trim(Replace(TxtObsA.Text, "'", "`"))
                RECUPERACION.R_IdLlamada = WS_CALL_ID
                RECUPERACION.R_Venta = 0
                RECUPERACION.R_Estado = IIf(GESTION.G_Intento >= CLIENTE.C_Intentos_Max, "T", "A")
                ' CLIENTE.CLI_ESTADO_OBJECION_CALIDAD = "2"
                RECUPERACION.R_Fecha_Vta = ""
                RECUPERACION.R_Hora_Vta = ""
                Dim fecha_agendamiento As String
                fecha_agendamiento = Format(DTAgenFecha2.Value, "dd/MM/yyyy") & " a las " & CmbHora2.Text & ":" & CmbMin2.Text & " Hrs."
                biGesRes.GuardaClienteGes(CLIENTE, RECUPERACION, GESTION)
                'biGesRes.ActualizaClienteAgen(CLIENTE.C_Id, CLIENTE.CLI_ESTADO_OBJECION_CALIDAD, CLIENTE.CLI_CALL_ID_CALIDAD)
                MsgBox("Fin de la gestión. Presione ACEPTAR para salir del formulario.", MsgBoxStyle.Information, csNombreAplicacion)
                limpiarPrimeraPantalla()
                Me.Hide()
                frmRegrabaciones.ShowDialog()
                BuscaGes()
            End If
        End If
    End Sub

    Private Sub TxtFechaN_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
        KeyAscii = CShort(SoloNumeros(KeyAscii))
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
        KeyAscii = CShort(SoloNumeros(KeyAscii))
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
        KeyAscii = CShort(SoloNumeros(KeyAscii))
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub

    '****************Metodo para que se cargue ciudad al seleccionar comuna de combobox cmbComuna************************
    Private Sub cmbComuna_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbComuna.SelectedIndexChanged
        Dim lstCiudad As New List(Of eCiudad)

        If cmbComuna.ValueMember Is Nothing Or cmbComuna.ValueMember = "" Then
            Exit Sub
        End If
        If cmbComuna.SelectedValue Is Nothing Then
            Exit Sub
        End If

        'actualiza el combo box de ciudad
        Dim Ciudad As New eCiudad
        Dim comuna As New eComuna

        comuna = vgListComuna.Find(Function(tmpC As eComuna) tmpC.idComuna = cmbComuna.SelectedValue)
        ' Ciudad = biClsCiudad.BuscaCiudadPorIdCiudad(comuna.idCiudad)
        Ciudad = vgListCiudad.Find(Function(tmc As eCiudad) tmc.idCiudad = comuna.idCiudad)
        lstCiudad.Add(Ciudad)
        vgFuncionComun.llenaComboBox(cmbCiudad, "nombreCiudad", "idCiudad", lstCiudad.ToArray)

    End Sub

    '**************metodo de boton siguiente de tab certificador********************************
    Private Sub cmdSiguienteFin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSiguienteFin.Click


        Select Case cmbAceptaPrima.SelectedIndex

            Case -1, 0
                MsgBox("Debe ingresar si acepta Cargo", MsgBoxStyle.Exclamation, csNombreAplicacion)
                cmbAceptaPrima.Focus()
                Exit Sub
            Case 1
                Select Case cmbAceptaContrato.SelectedIndex
                    Case 0, -1
                        MsgBox("Debe ingresar si acepta Contratación", MsgBoxStyle.Exclamation, csNombreAplicacion)
                        cmbAceptaContrato.Focus()
                        Exit Sub
                    Case 1
                        If perfil <> "Regrabador" Then
                            ' CLIENTE.cli_acepta_cargo = cmbAceptaPrima.Text
                            ' CLIENTE.cli_acepta_contrato = cmbAceptaContrato.Text
                        Else
                            'CLIENTE.CLI_ACEPTA_CONTRATACION = cmbAceptaContrato.Text
                            'CLIENTE.cli_acepta_prima = cmbAceptaPrima.Text
                        End If

                        clsScript = CargaScript(6)
                        WebBrowsercierre.DocumentText = clsScript.contenidoScript

                        Cuerpo.TabPages.Add(_Cuerpo_UltInfo)
                        Cuerpo.TabPages.Item(0).Parent = Nothing
                        guardarPantallaAnterior(6)
                    Case 2
                        If perfil <> "Regrabador" Then
                            ' CLIENTE.cli_acepta_contrato = cmbAceptaContrato.Text
                            ' CLIENTE.cli_acepta_prima = cmbAceptaPrima.Text
                        Else
                            ' CLIENTE.CLI_ACEPTA_CONTRATACION = cmbAceptaContrato.Text
                            ' CLIENTE.cli_acepta_prima = cmbAceptaPrima.Text
                        End If

                        Cuerpo.TabPages.Add(_Cuerpo_Objeciones)
                        Cuerpo.TabPages.Item(0).Parent = Nothing
                        guardarPantallaAnterior(6)

                End Select

            Case 2
                ' CLIENTE.cli_acepta_contrato = cmbAceptaContrato.Text
                ' CLIENTE.cli_acepta_prima = cmbAceptaPrima.Text
                Cuerpo.TabPages.Add(_Cuerpo_Objeciones)
                Cuerpo.TabPages.Item(0).Parent = Nothing
                guardarPantallaAnterior(6)
        End Select

    End Sub

    Private Sub InsertaAdicionales()
        Dim J As Integer
        Dim biAsegurado As New clsAdicionalBI
        Dim adi As New eAdicional

        biAsegurado.Eliminar(CLIENTE.C_Id)

        For J = 0 To lstAdi.Count - 1
            adi.C_ID = CLIENTE.C_Id
            adi.A_NRO = J + 1
            adi.A_Rut = IIf(lstAdi(J).A_Rut.Trim = "", 0, lstAdi(J).A_Rut)
            adi.A_Dv = lstAdi(J).A_Dv
            adi.A_Nombre = lstAdi(J).A_Nombre
            adi.A_Paterno = lstAdi(J).A_Paterno
            adi.A_Materno = lstAdi(J).A_Materno
            adi.A_FechaNacimiento = Convert.ToDateTime(lstAdi(J).A_FechaNacimiento).ToString("yyyyMMdd")
            adi.A_Sexo = lstAdi(J).A_Sexo
            adi.A_TipoBeneficiario = lstAdi(J).A_TipoBeneficiario
            adi.A_NombreBeneficiario = lstAdi(J).A_NombreBeneficiario
            adi.a_primaUf = lstAdi(J).a_primaUf
            adi.a_primaPesos = lstAdi(J).a_primaPesos
            adi.idPlanAdic = lstAdi(J).idPlanAdic
            adi.a_salud = lstAdi(J).a_salud
            biAsegurado.Insertar(adi)
        Next J

    End Sub

    Private Sub InsertaBeneficiarios()
        Dim J As Integer
        Dim clsBen As New clsBeneficiarioBI
        Dim be As New eBeneficiario

        clsBen.Eliminar(CLIENTE.C_Id)
        For J = 0 To lstBen.Count - 1
            be.C_ID = CLIENTE.C_Id
            be.B_Nro = J + 1
            be.B_Rut = IIf(lstBen(J).B_Rut.Trim = "", 0, lstBen(J).B_Rut)
            be.B_Dv = lstBen(J).B_Dv
            be.B_Nombre = lstBen(J).B_Nombre
            be.B_Paterno = lstBen(J).B_Paterno
            be.B_Materno = lstBen(J).B_Materno
            be.B_TipoBeneficiario = lstBen(J).B_TipoBeneficiario
            be.B_NOMBRE_PARENTESCO = lstBen(J).B_NOMBRE_PARENTESCO
            be.B_Porcentaje = lstBen(J).B_Porcentaje
            be.B_FechaNacimiento = lstBen(J).B_FechaNacimiento
            be.B_Sexo = lstBen(J).B_Sexo
            clsBen.Insertar(be)
        Next J

    End Sub

#Region "INGRESO CLIENTE"
    '*********metodo para cargar los datos en la toma de datos de regrabacion*******************
    ''' <summary>
    ''' carga los datos a los controles del tab toma de datos
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AsignadatosCli()
        Dim TotalDigito As Integer

        TotalDigito = Len(Fono_A_Llamar)
        If TotalDigito = 9 Then
            txtFonoVenta.Text = Trim(Fono_A_Llamar)
        End If
        If TotalDigito > 9 Then
            txtFonoVenta.Text = Microsoft.VisualBasic.Right(Fono_A_Llamar, 9)
        End If

        'vgFuncionComun.LimpiaFormilarioBenAdi()

        Dim nombre = ""
        Dim apellido = ""
        Dim index As Integer = Trim(CLIENTE.C_Nombre).IndexOf(" "c)
        If (index = -1) Then
            ' No existe ningún espacio en blanco;
            nombre = Trim(CLIENTE.C_Nombre)
            apellido = String.Empty

        Else
            ' Obtenemos el nombre
            nombre = Trim(CLIENTE.C_Nombre).Substring(0, index)

            ' Obtenemos el apellido
            apellido = Trim(CLIENTE.C_Nombre).Substring(index + 1, Trim(CLIENTE.C_Nombre).Length - index - 1)

        End If

        txtNombreV.Text = nombre
        txtPaternoV.Text = Trim(UCase(CLIENTE.C_Paterno))
        txtMaternoV.Text = Trim(UCase(CLIENTE.C_Materno))

        txtRutV.Text = Mid(CLIENTE.C_Rut, 1, 4)
        txtDvV.Text = ""
        txtEmail.Text = CLIENTE.C_Email
        txtCalleV.Text = "" 'CLIENTE.c_direccion
        txtUltDigitos.Text = ""

        'DatosMedioPago()
        lstAdi.Clear()

    End Sub

    Private Sub DatosMedioPago()
        Dim biMedioPago As New clsMedioPagoBI
        Dim listMedioPago As New List(Of eMedioPago)
        Dim medioPagoVacio As New eMedioPago
        Dim listMedioPagoVacio As New List(Of eMedioPago)

        medioPagoVacio.MedioPago = "[No Especificado]"
        medioPagoVacio.idMedioPago = "0"
        listMedioPagoVacio.Add(medioPagoVacio)
        listMedioPago = biMedioPago.BuscaDatosMedioPago(vgCampania.idCRM, CLIENTE.C_Id)

        For x As Int16 = 0 To listMedioPago.Count - 1

            Dim medioPago As New eMedioPago
            medioPago.MedioPago = listMedioPago(x).MedioPago
            medioPago.NumeroMedioPago = listMedioPago(x).NumeroMedioPago
            medioPago.idMedioPago = listMedioPago(x).idMedioPago
            medioPago.otroMedioPago = listMedioPago(x).otroMedioPago

            If (medioPago.otroMedioPago) Then
                medioPago.MedioPago = listMedioPago(x).Tarjeta
            Else
                medioPago.MedioPago = listMedioPago(x).Tarjeta & " - Nro: " & medioPago.NumeroMedioPago
            End If

            listMedioPagoVacio.Add(medioPago)

        Next

        vgFuncionComun.llenaComboBox(cmbMedioPago, "MedioPago", "idMedioPago", listMedioPagoVacio.ToArray)

        If perfil = "Regrabador" Then
            cmbMedioPago.Text = GESTION.G_MEDIO_PAGO
            If GESTION.G_BANCO <> "" Then
                CmbBanco.Text = GESTION.G_BANCO
                txtNumeroCta.Text = GESTION.G_NUMERO_TARJETA
                Dim vencimiento As String() = GESTION.G_VENCIMIENTO_TARJETA.Split("/")
                cmbMes.Text = vencimiento(0)
                cmbAnio.Text = vencimiento(1)
                cmbTpoTarjeta.Text = GESTION.G_TIPO_TARJETA
                cmbDiaPago.Text = GESTION.G_DIA_PAGO
            End If
        End If

    End Sub

    ''' <summary>
    ''' metodo para cargar los datos en la toma de datos de regrabacion
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AsignadatosCliGes()
        Dim comunaGes As New eComuna
        Dim CiudadGes As New eCiudad
        Dim tipocontrato As New eTipoContrato
        Dim biciudad As New clsCiudadBI
        Dim lstCiudad As New List(Of eCiudad)
        Dim plan As New ePlan

        txtNombreV.Text = Trim(UCase(GESTION.G_Nombre))
        txtPaternoV.Text = Trim(UCase(GESTION.G_Paterno))
        txtMaternoV.Text = Trim(UCase(GESTION.G_Materno))

        txtFonoVenta.Text = Trim(UCase(GESTION.G_Fono_Contacto))
        dtFechaNacV.Text = Trim(UCase(GESTION.G_Fecha_Nacimiento))
        CmPrevision.Text = GESTION.G_Dato1.ToUpper

        txtRutV.Text = Trim(GESTION.G_Rut)
        txtDvV.Text = Trim(GESTION.G_Dv)
        txtEmail.Text = Trim(GESTION.G_Email)
        txtFonoVenta.Text = Trim(GESTION.G_Fono_Contacto)
        txtCalleV.Text = Trim(GESTION.G_Calle)
        txtNroV.Text = Trim(GESTION.G_Nro)
        txtReferenciaV.Text = Trim(GESTION.G_Referencia)


        comunaGes = vgListComuna.Find(Function(tmpC As eComuna) tmpC.nombreComuna = GESTION.G_Comuna)
        cmbComuna.SelectedValue = comunaGes.idComuna
        CiudadGes = biciudad.BuscaCiudadPorIdCiudad(comunaGes.idCiudad)
        lstCiudad.Add(CiudadGes)
        vgFuncionComun.llenaComboBox(cmbCiudad, "nombreCiudad", "idCiudad", lstCiudad.ToArray)
        cmbCiudad.SelectedValue = CiudadGes.idCiudad


        If GESTION.G_Sexo = "M" Then
            CmbSexo.SelectedIndex = 1
        ElseIf GESTION.G_Sexo = "F" Then
            CmbSexo.SelectedIndex = 2
        Else
            CmbSexo.SelectedIndex = 0
        End If

        llenarTipoContrato()
        tipocontrato = vgListTipoContrato.Find(Function(tmpC As eTipoContrato) tmpC.idTipoContrato = GESTION.G_TipoContrato)
        cmbTipoContrato.SelectedValue = tipocontrato.idTipoContrato

        llenar_planes()

        plan = vgListPlanes.Find(Function(tmpC As ePlan) tmpC.idPlan = GESTION.G_Plan)
        cmbPlan.SelectedValue = plan.idPlan

        cmbTipoContrato.Enabled = False
        cmbPlan.Enabled = False

        Carga_Beneficiarios()
        Carga_adicionales()

    End Sub


    Private Sub limpiarPrimeraPantalla()
        TxtRut.Text = ""
        txtNombre.Text = ""
        TxtId.Text = ""
        txtDireccion.Text = ""
        txtIntentos.Text = ""
        txtObsAgen.Text = ""
        CmdAtras.Enabled = False
        lstBen.Clear()
        'limpiaring()

        frmBeneficiario.dtBeneficiario.DataSource = Nothing
        frmBeneficiario.dtBeneficiario.Rows.Clear()

    End Sub




#End Region

#Region "VALIDACIONES"
    Function DevuelveDias(ByVal fechaTermino As Date) As Integer
        Dim cantDias As Integer
        Dim i As Integer
        Dim j As Integer
        Dim diaActual As Date
        Dim CantFeriados As Integer
        CantFeriados = 0
        cantDias = 0
        fechaTermino = fechaTermino.Date()
        cantDias = DateDiff(DateInterval.Day, Now.Date, fechaTermino)
        diaActual = Now.Date

        For i = 0 To cantDias
            For j = 0 To UBound(feriados)
                If diaActual = feriados(j) Then
                    CantFeriados = CantFeriados + 1
                End If
            Next
            diaActual = DateAdd(DateInterval.Day, 1, diaActual)
        Next

        DevuelveDias = cantDias - CantFeriados + 1
    End Function

    ''' <summary>
    ''' validamos la informacion en la toma de datos
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ValidaInformacion() As Boolean

        ValidaInformacion = False

        '' Validar Nombre
        If Len(Trim(txtNombreV.Text)) < 3 Or
            Len(Trim(txtPaternoV.Text)) < 3 Or
            Len(Trim(txtMaternoV.Text)) < 3 Then
            MsgBox("Debe ingresar correctamente el NOMBRE COMPLETO del cliente ", vbInformation, "CORDIALPHONE.")
            txtNombreV.Focus()
            Exit Function
        End If

        If Not vgFuncionComun.edad(dtFechaNacV.Value, csEdadMinima, csEdadMaxima) Then
            MsgBox("La fecha de nacimiento no es válida.", vbInformation, "CORDIALPHONE.")
            Exit Function

        End If

        Select Case CmbSexo.SelectedIndex
            Case 0, -1
                MsgBox("Debe ingresar sexo del cliente.", MsgBoxStyle.Exclamation, csNombreAplicacion)
                CmbSexo.Focus()
                Exit Function
        End Select

        If txtRutV.Text.Length < 7 Then
            MsgBox("El RUT del cliente no es valido.", vbInformation, "Callsotuh.")
            txtRutV.Focus()
            Exit Function
        End If

        If Not vgFuncionComun.validarRut(Trim(txtRutV.Text) & "-" & Trim(txtDvV.Text)) Then
            MsgBox("El RUT del cliente no es valido.", vbInformation, "CORDIALPHONE.")
            txtRutV.Focus()
            Exit Function
        End If

        'If perfil = "Regrabador" Then
        '    If txtRutV.Text & txtDvV.Text.ToUpper <> GESTION.G_Rut & GESTION.G_Dv.ToUpper Then
        '        MsgBox("El RUT del cliente es distinto al de la base original.", vbInformation, "CORDIALPHONE.")
        '        txtRutV.Focus()
        '        Exit Function
        '    End If
        'Else
        '    If txtRutV.Text & txtDvV.Text.ToUpper <> CLIENTE.C_Rut & CLIENTE.C_Dv.ToUpper Then
        '        MsgBox("El RUT del cliente es distinto al de la base original.", vbInformation, "CORDIALPHONE.")
        '        txtRutV.Focus()
        '        Exit Function
        '    End If
        'End If
        If txtFonoVenta.Text = "" Or txtFonoVenta.TextLength < 9 Then
            MsgBox("Debe ingresar fono venta valido 9 dígitos.", vbExclamation, csNombreAplicacion)
            txtFonoVenta.Focus()
            Exit Function
        End If


        'If Trim(txtEmail.Text) <> "" Then
        '    If vgFuncionComun.ValidaEmail(txtEmail.Text) = False Then
        '        MsgBox("Correo ingresado no es valido.", MsgBoxStyle.Exclamation, csNombreAplicacion)
        '        txtEmail.Focus()
        '        Exit Function
        '    End If
        'Else
        '    MsgBox("Debe ingresar un email", vbExclamation, csNombreAplicacion)
        '    txtEmail.Focus()
        '    Exit Function
        'End If

        If CmPrevision.SelectedIndex < 1 Then
            MsgBox("Debe seleccionar la previsión del cliente", vbExclamation, csNombreAplicacion)
            CmPrevision.Focus()
            Exit Function
        End If

        If Trim(txtEmail.Text) <> "" Then
            If vgFuncionComun.ValidaEmail(Trim(txtEmail.Text)) = False Then
                MsgBox("Correo ingresado no es valido.", MsgBoxStyle.Exclamation, csNombreAplicacion)
                txtEmail.Focus()
                Exit Function
            End If
        End If

        Dim correo As New eCorreoInvalido
        If txtEmail.Text = "" Then
            If MsgBox("Pasara la venta sin correo, desea proseguir?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then Exit Function
        Else

            'correo = listaCorreoInvalido.Find(Function(x As eCorreoInvalido) x.correo = txtEmail.Text)
            'If Not correo Is Nothing Then
            '    MsgBox("El correo ingresado " + txtEmail.Text + " NO es válido, se eliminará el correo al momento de guardar.", MsgBoxStyle.Information)
            '    txtEmail.Text = ""
            '    Exit Function
            'End If

        End If


        Select Case CmbSexo.SelectedIndex
            Case 0, -1
                MsgBox("Debe ingresar sexo del cliente.", MsgBoxStyle.Exclamation, csNombreAplicacion)
                CmbSexo.Focus()
                Exit Function
        End Select

        If Trim(txtCalleV.Text) = "" Then
            MsgBox("El Campo calle es Obligatorio", vbExclamation, csNombreAplicacion)
            txtCalleV.Focus()
            Exit Function
        End If

        If Trim(txtNroV.Text) = "" Then
            MsgBox("El Campo Nº es obligatorio, si no tiene número ingrese S/N", vbExclamation, csNombreAplicacion)
            txtNroV.Focus()
            Exit Function
        End If

        If InStr(1, Trim(txtCalleV.Text), Trim(txtNroV.Text)) > 0 Then
            If MsgBox("El Campo Nº ya se encuentra en la calle, ¿Se debe corregir esto? " & vbNewLine & txtCalleV.Text & " " & txtNroV.Text, vbOKCancel, "Validacion de Nro en Direccion") = vbOK Then
                txtCalleV.Focus()
                txtCalleV.SelectionStart = InStr(1, Trim(txtCalleV.Text), Trim(txtNroV.Text)) - 1
                txtCalleV.SelectionLength = Len(Trim(txtNroV.Text))
                Exit Function
            End If
        End If

        If cmbComuna.SelectedIndex = -1 Or cmbComuna.Text = "" Then
            MsgBox("Ingrese la comuna a la cual pertenece la dirección.", vbExclamation, csNombreAplicacion)
            cmbComuna.Focus()
            Exit Function
        End If


        If cmbCiudad.SelectedIndex = -1 Or cmbCiudad.Text = "" Or cmbCiudad.SelectedValue = 0 Then
            MsgBox("Ingrese la ciudad a la cual pertenece la comuna.", vbExclamation, csNombreAplicacion)
            cmbCiudad.Focus()
            Exit Function
        End If

        If cmbTipoContrato.SelectedIndex = -1 Or cmbTipoContrato.SelectedIndex = 0 Then
            MsgBox("Debe ingresar el tipo de contrato que se contratará.", vbExclamation, csNombreAplicacion)
            cmbTipoContrato.Focus()
            Exit Function
        End If

        If perfil = "Regrabador" Then
            If cmbPlan.SelectedIndex = -1 Then
                MsgBox("Debe selecionar el tipo de plan.", vbExclamation, csNombreAplicacion)
                cmbPlan.Focus()
                Exit Function
            End If
        Else
            If cmbPlan.SelectedValue = 0 Then
                MsgBox("Debe selecionar el tipo de plan.", vbExclamation, csNombreAplicacion)
                cmbPlan.Focus()
                Exit Function
            End If
        End If
        ValidaInformacion = True
    End Function

#End Region

#Region "ADICIONALES"
    ''' <summary>
    ''' PROCEDIMIENTO PARA CARGAR LOS ADICIONALES EN LA GRILLA
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Carga_adicionales()
        lstAdi = biClsAdic.Carga_Adicional(CLIENTE.C_Id)
        sumaUFAdicionales()
    End Sub

    Private Sub dtAdicional_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dtAdicional.CellClick

        If e.RowIndex = -1 Then Exit Sub

        txtNombreA.Text = dtAdicional.Rows(e.RowIndex).Cells("nombre").Value
        txtPaternoA.Text = dtAdicional.Rows(e.RowIndex).Cells("paterno").Value
        txtMaternoA.Text = dtAdicional.Rows(e.RowIndex).Cells("materno").Value
        txtRutA.Text = dtAdicional.Rows(e.RowIndex).Cells("rut").Value
        txtDvA.Text = dtAdicional.Rows(e.RowIndex).Cells("dv").Value
        cmbParentescoAdic.Text = dtAdicional.Rows(e.RowIndex).Cells("tipo_parentesco").Value
        dtFechaNacAdic.Value = dtAdicional.Rows(e.RowIndex).Cells("fechaNacimiento").Value
        cmbSexoA.Text = dtAdicional.Rows(e.RowIndex).Cells("Sexo").Value

        numeroFila = e.RowIndex
    End Sub

    Private Sub btnAdicional_Click(sender As Object, e As EventArgs) Handles btnAdicional.Click
        frmAdicionales.ShowDialog()
        sumaUFAdicionales()
    End Sub


#End Region

#Region "BENEFICIARIOS"
    Private Sub Carga_Beneficiarios()
        lstBen = biClsBen.CargaBeneficiarios(CLIENTE.C_Id)
    End Sub

    Private Sub btnBeneficiarios_Click(sender As Object, e As EventArgs) Handles btnBeneficiarios.Click
        If cmbTipoContrato.SelectedIndex <> 0 And cmbTipoContrato.SelectedIndex <> -1 Then

            frmBeneficiario.ShowDialog()
            'Me.Close()

        End If
    End Sub

#End Region

#Region "SCRIPTS"

    Private Function ScriptLblFinNoC() As String
        Dim mensaje As String
        mensaje = "Disculpe las molestias, gracias por su tiempo y que tenga un buen día." & vbNewLine
        ScriptLblFinNoC = mensaje
    End Function

    Private Function ScriptLblFinNoCumple() As String
        Dim mensaje As String
        mensaje = "Sr./Sra " & Trim$(CLIENTE.C_Nombre) & "  " & Trim$(CLIENTE.C_Paterno) & " " & Trim$(CLIENTE.C_Materno) & " Lamentablemente no podemos entregar el seguro ya no cumple con alguno de los requisitos anteriormente mencionado. Muchas gracias por su tiempo. " & vbNewLine
        ScriptLblFinNoCumple = mensaje
    End Function

#End Region


    Public Function CargaScript(ByVal _idTipoScript As Int32) As eScript

        Dim script As New eScript
        Dim biScript As New clsScriptBI
        script = biScript.BuscarScriptPorIdTipoScript(vgCampania.idCRM, _idTipoScript)

        'datos generales
        script.contenidoScript = Replace(script.contenidoScript, "[FechaActual]", Now())
        script.contenidoScript = Replace(script.contenidoScript, "[Persona.Campo1]", CLIENTE.Extra1)
        script.contenidoScript = Replace(script.contenidoScript, "[Persona.Campo2]", CLIENTE.Extra2)
        script.contenidoScript = Replace(script.contenidoScript, "[Persona.Campo3]", CLIENTE.Extra3)
        script.contenidoScript = Replace(script.contenidoScript, "[Persona.Campo4]", CLIENTE.Extra4)
        script.contenidoScript = Replace(script.contenidoScript, "[Persona.Campo5]", CLIENTE.Extra5)
        script.contenidoScript = Replace(script.contenidoScript, "[Persona.Campo6]", CLIENTE.Extra6)
        script.contenidoScript = Replace(script.contenidoScript, "[Persona.Campo7]", CLIENTE.Extra7)
        script.contenidoScript = Replace(script.contenidoScript, "[Persona.Campo8]", CLIENTE.Extra8)
        script.contenidoScript = Replace(script.contenidoScript, "[Persona.Campo9]", CLIENTE.Extra9)
        script.contenidoScript = Replace(script.contenidoScript, "[Persona.Campo10]", CLIENTE.Extra10)

        If (script.idTipoScript < 3) Then '1.Bienvenida  2.Presentacion  3.Informacion Adicional

            '[Persona.nombre] [Persona.paterno] [Persona.materno]
            script.contenidoScript = Replace(script.contenidoScript, "[Agente.Nombre]", Replace(WS_IDUSUARIO, ".", " "))
            script.contenidoScript = Replace(script.contenidoScript, "[Persona.nombre]", CLIENTE.C_Nombre)
            script.contenidoScript = Replace(script.contenidoScript, "[Persona.paterno]", CLIENTE.C_Paterno)
            script.contenidoScript = Replace(script.contenidoScript, "[Persona.materno]", CLIENTE.C_Materno)
            script.contenidoScript = Replace(script.contenidoScript, "[Persona.fechaNacimiento]", CLIENTE.C_Fecha_Nacimiento)
            script.contenidoScript = Replace(script.contenidoScript, "[Persona.Rut]", CLIENTE.C_Rut & "-" & CLIENTE.C_Dv)
            script.contenidoScript = Replace(script.contenidoScript, "[DireccionParticular]", CLIENTE.C_Direccion)
            script.contenidoScript = Replace(script.contenidoScript, "[Persona.Comuna]", CLIENTE.C_Comuna)
            script.contenidoScript = Replace(script.contenidoScript, "[Persona.Ciudad]", CLIENTE.C_Ciudad)

        Else
            'datos venta

            script.contenidoScript = Replace(script.contenidoScript, "[Persona.nombre]", GESTION.G_Nombre)
            script.contenidoScript = Replace(script.contenidoScript, "[Persona.paterno]", GESTION.G_Paterno)
            script.contenidoScript = Replace(script.contenidoScript, "[Persona.materno]", GESTION.G_Materno)
            script.contenidoScript = Replace(script.contenidoScript, "[Persona.fechaNacimiento]", GESTION.G_Fecha_Nacimiento)
            script.contenidoScript = Replace(script.contenidoScript, "[Persona.Rut]", GESTION.G_Rut & "-" & GESTION.G_Dv)
            script.contenidoScript = Replace(script.contenidoScript, "[DireccionParticular]", GESTION.G_Calle & " " & GESTION.G_Nro & " " & GESTION.G_Referencia)
            script.contenidoScript = Replace(script.contenidoScript, "[Persona.Comuna]", GESTION.G_Comuna)
            script.contenidoScript = Replace(script.contenidoScript, "[Persona.mail]", GESTION.G_Email)
            script.contenidoScript = Replace(script.contenidoScript, "[Persona.FonoVenta]", GESTION.G_Fono_Contacto)
            script.contenidoScript = Replace(script.contenidoScript, "[Persona.FonoContacto]", GESTION.G_Fono_Contacto)
            'script.contenidoScript = Replace(script.contenidoScript, "[medioPago]", GESTION.G_ME)
            'script.contenidoScript = Replace(script.contenidoScript, "[banco]", CLIENTE.CLI_ABANCO)
            script.contenidoScript = Replace(script.contenidoScript, "[PrimaUf]", GESTION.G_Prima_Uf)
            script.contenidoScript = Replace(script.contenidoScript, "[PrimaPesos]", GESTION.G_Prima_Pesos)
            'script.contenidoScript = Replace(script.contenidoScript, "[Persona.numeroTarjeta]", CLIENTE.cli_anrotarjeta)
            script.contenidoScript = Replace(script.contenidoScript, "[Persona.email]", GESTION.G_Email)
            'script.contenidoScript = Replace(script.contenidoScript, "[CodigoVerificacion]", CLIENTE.cli_codverificacion)

            'Beneficiarios

            For x As Int16 = 0 To lstBen.Count - 1

                script.contenidoScript = Replace(script.contenidoScript, "BeneficiarioNombre" + (x + 1).ToString, lstBen(x).B_Nombre)
                script.contenidoScript = Replace(script.contenidoScript, "BeneficiarioPaterno" + (x + 1).ToString, lstBen(x).B_Paterno)
                script.contenidoScript = Replace(script.contenidoScript, "BeneficiarioMaterno" + (x + 1).ToString, lstBen(x).B_Materno)
                script.contenidoScript = Replace(script.contenidoScript, "BeneficiarioFechaNacimiento" + (x + 1).ToString, lstBen(x).B_FechaNacimiento)
                script.contenidoScript = Replace(script.contenidoScript, "BeneficiarioRut" + (x + 1).ToString, lstBen(x).B_Rut + "-" + lstBen(x).B_Dv)
                script.contenidoScript = Replace(script.contenidoScript, "BeneficiarioParentesco" + (x + 1).ToString, lstBen(x).B_TipoBeneficiario)
                script.contenidoScript = Replace(script.contenidoScript, "BeneficiarioPorcentaje" + (x + 1).ToString, lstBen(x).B_Porcentaje)

            Next

        End If

        Return script

    End Function

    Private Sub llenar_planes()
        Dim biplan As New clsPlanBI

        vgListPlanes = biplan.BuscarPlanPorTipoContrato(cmbTipoContrato.SelectedValue, vgCampania.idCRM)

        If vgListPlanes.Count > 1 Then
            vgFuncionComun.llenaComboBox(cmbPlan, "descripcionPlan", "idPlan", vgListPlanes.ToArray)
            cmbPlan.Visible = True
            lblPlanes.Visible = True
        Else
            cmbPlan.Visible = False
            lblPlanes.Visible = False
        End If
    End Sub

    Private Sub cmbPlan_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbPlan.SelectedIndexChanged
        If cmbPlan.SelectedIndex <> -1 And cmbPlan.SelectedIndex <> 0 Then
            sumaUFAdicionales()
            RemoveHandler cmbPlan.SelectedIndexChanged, AddressOf cmbPlan_SelectedIndexChanged
            AddHandler cmbPlan.SelectedIndexChanged, AddressOf cmbPlan_SelectedIndexChanged
        Else
            lblPrimaUF.Text = 0
            lblPrimaPesos.Text = 0
        End If
    End Sub

    Private Sub dtFechaNac_ValueChanged(sender As Object, e As EventArgs) Handles dtFechaNacV.ValueChanged

        If dtFechaNacV.Text <> "" Then
            txtCalculaEdad.Text = DateDiff(DateInterval.Year, Date.Parse(dtFechaNacV.Text), Date.UtcNow)
        End If

        If cmbTipoContrato.SelectedIndex <> -1 And cmbTipoContrato.SelectedIndex <> 0 Then
            LlenaPlanesCondicion()
            vgFuncionComun.sumaUFAdicionales()
        Else
            lblPrimaUF.Text = "0"
            lblPrimaPesos.Text = "0"
        End If

        cmbPlan.Enabled = True
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim biMedioPago As New clsMedioPagoBI

        If validaMedioPago() Then
            GESTION.G_BANCO = IIf(medioPagoGlobal.otroMedioPago, CmbBanco.Text, "")
            GESTION.G_NUMERO_TARJETA = IIf(medioPagoGlobal.otroMedioPago, txtNumeroCta.Text, txtUltDigitos.Text)
            GESTION.G_MEDIO_PAGO = cmbMedioPago.Text

            If cmbMes.SelectedIndex >= 1 Then
                GESTION.G_VENCIMIENTO_TARJETA = IIf(medioPagoGlobal.otroMedioPago, cmbMes.Text + "/" + cmbAnio.Text, 0)
                GESTION.G_TIPO_TARJETA = IIf(medioPagoGlobal.otroMedioPago, cmbTpoTarjeta.Text, "")
            Else
                GESTION.G_VENCIMIENTO_TARJETA = ""
                GESTION.G_TIPO_TARJETA = ""
            End If

            'txtNumeroCta.Text
            'CLIENTE.CLI_DIACARGO = IIf(CmbDiaCargo.Visible = True, CmbDiaCargo.Text, "0")

            'CERTIFIOCACION
            clsScript = CargaScript(4)
            WebInfAdicional.DocumentText = clsScript.contenidoScript

            Cuerpo.TabPages.Add(_Cuerpo_InfAdic)
            Cuerpo.TabPages.Item(0).Parent = Nothing
            guardarPantallaAnterior(4)

        End If

    End Sub

    Private Function validaMedioPago() As Boolean
        validaMedioPago = False
        If medioPagoGlobal.otroMedioPago Then
            If CmbBanco.SelectedIndex < 1 Then
                MsgBox("Debe ingresar el banco.", MsgBoxStyle.Information, "CORDIALPHONE")
                CmbBanco.Focus()
                Exit Function
            End If

            If Not CmbBanco.Text.Trim() = "BANCO DE CREDITO E INVERSIONES" Then
                If txtNumeroCta.Text.Trim = "" Or Len(txtNumeroCta.Text.Trim) < 16 Then
                    MsgBox("Debe ingresar los números de la tarjeta.", MsgBoxStyle.Information, "CORDIALPHONE")
                    txtNumeroCta.Focus()
                    Exit Function
                End If

                If cmbMes.SelectedIndex < 1 Then
                    MsgBox("Debe ingresar el mes vencimiento de tarjeta.", MsgBoxStyle.Information, "CORDIALPHONE")
                    cmbMes.Focus()
                    Exit Function
                End If

                If cmbAnio.SelectedIndex < 1 Then
                    MsgBox("Debe ingresar el mes vencimiento de tarjeta.", MsgBoxStyle.Information, "CORDIALPHONE")
                    cmbAnio.Focus()
                    Exit Function
                End If

                If cmbTpoTarjeta.SelectedIndex < 1 Then
                    MsgBox("Debe seleccionar el tipo de tarjeta.", MsgBoxStyle.Information, "CORDIALPHONE")
                    cmbTpoTarjeta.Focus()
                    Exit Function
                End If
            Else
                If txtNumeroCta.Text.Trim = "" Or Len(txtNumeroCta.Text.Trim) < 3 Then
                    MsgBox("Debe ingresar los números de la tarjeta.", MsgBoxStyle.Information, "CORDIALPHONE")
                    txtNumeroCta.Focus()
                    Exit Function
                End If
            End If
        Else
            If cmbMedioPago.SelectedIndex < 1 Then
                MsgBox("Debe ingresar el medio de pago seleccionado por el cliente.", MsgBoxStyle.Information, "CORDIALPHONE")
                cmbMedioPago.Focus()
                Exit Function
            End If

            If txtUltDigitos.Text.Length > 4 Then
                MsgBox("Debe ingresar los 3 o 4 ultimos números.", MsgBoxStyle.Information, "CORDIALPHONE")
                txtUltDigitos.Focus()
                Exit Function
            End If

        End If
        validaMedioPago = True
    End Function

    Private Sub cmbAceptaPrima_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbAceptaPrima.SelectedIndexChanged
        If cmbAceptaPrima.SelectedIndex = 1 Then

            'Script de certificación
            clsScript = CargaScript(5)
            WebBrowsercierre.DocumentText = clsScript.contenidoScript

            WebBrowsercierre.Visible = True
        Else
            WebBrowsercierre.Visible = False
        End If
    End Sub

    Private Sub cmbTipoContrato_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbTipoContrato.SelectedIndexChanged

        If cmbTipoContrato.SelectedIndex <> -1 And cmbTipoContrato.SelectedIndex <> 0 Then
            llenar_planes()

            If dtFechaNacV.Text <> "" Then
                txtCalculaEdad.Text = DateDiff(DateInterval.Year, Date.Parse(dtFechaNacV.Text), Date.UtcNow)
            End If

            LlenaPlanesCondicion()

            If vgListTipoContrato.Item(cmbTipoContrato.SelectedIndex).cantidadBeneficiarios <> 0 Then
                btnBeneficiarios.Visible = True
            Else
                btnBeneficiarios.Visible = False
            End If

            If vgListTipoContrato.Item(cmbTipoContrato.SelectedIndex).cantidadAdicionales <> 0 Then
                btnAdicional.Visible = True
            Else
                btnAdicional.Visible = False
            End If
        Else
            lblPlanes.Visible = False
            cmbPlan.Visible = False
            lblPrimaUF.Text = "0"
            lblPrimaPesos.Text = "0"
            btnAdicional.Visible = False
            btnBeneficiarios.Visible = False
        End If
    End Sub

    Private Sub LlenaPlanesCondicion()
        Dim tmpPlanes As New List(Of ePlan)
        Dim plan As New ePlan
        plan.idPlan = 0
        plan.primaUF = "---No ingresado---"
        plan.descripcionPlan = "---No ingresado---"
        tmpPlanes.Add(plan)

        If cmbTipoContrato.SelectedIndex <> -1 And cmbTipoContrato.SelectedIndex <> 0 Then

            tipoContrato.idTipoContrato = cmbTipoContrato.SelectedValue 'siempre analizara restriccion del titular
            listPlanes = biClsPlan.BuscarPlanPorTipoContrato(tipoContrato.idTipoContrato, vgCampania.idCRM)
            Dim listRestriccion As New List(Of eRestriccion)

            For x As SByte = 0 To listPlanes.Count - 1
                If listPlanes(x).idPlan = 0 Then
                    Continue For
                End If
                planE = biClsPlan.BuscarPlanPorIdPlan(listPlanes(x).idPlan)

                listRestriccion = biClsRestricion.BuscarRestriccionPorIdPlan(planE.idPlan)
                Dim count As Int16 = 0
                For y As SByte = 0 To listRestriccion.Count - 1
                    restricionE.idPlanesPorEdad = listRestriccion(y).idPlanesPorEdad
                    restricionE.idPlan = listRestriccion(y).idPlan
                    restricionE.edadDesde = listRestriccion(y).edadDesde
                    restricionE.edadHasta = listRestriccion(y).edadHasta

                    Dim edadCliente As Int16 = edad(dtFechaNacV.Value) 'DateDiff(DateInterval.Year, Date.Parse(dtFechaNac.Text), Date.UtcNow)

                    Select Case edadCliente
                        Case restricionE.edadDesde To restricionE.edadHasta
                            count = 1
                    End Select
                Next

                If count = 1 Then
                    tmpPlanes.Add(planE)
                    Dim TpoContrato As New eTipoContrato
                    TpoContrato = biClsTipoContrato.BuscarTipoContratoPorIdTipoContrato(tipoContrato.idTipoContrato)
                    btnAdicional.Visible = IIf(TpoContrato.cantidadAdicionales > 0, True, False)
                    btnBeneficiarios.Visible = IIf(TpoContrato.cantidadBeneficiarios > 0, True, False)
                End If
            Next
        End If

        If tmpPlanes.ToArray.Count > 0 Then
            cmbPlan.Visible = True
            lblPlanes.Visible = True
            vgFuncionComun.llenaComboBox(cmbPlan, "descripcionPlan", "idplan", tmpPlanes.ToArray)
        Else
            cmbPlan.Visible = False
            lblPlanes.Visible = False
        End If

    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        clsScript = CargaScript(5)
        wbScriptCertificacion.DocumentText = clsScript.contenidoScript

        Cuerpo.TabPages.Add(_Cuerpo_Certifica)
        Cuerpo.TabPages.Item(0).Parent = Nothing
        guardarPantallaAnterior(5)
    End Sub

    Private Sub CmdFinVenta_Click(sender As Object, e As EventArgs) Handles CmdFinVenta.Click

        GESTION.G_Estado = "T"
        GESTION.G_Venta = "1"
        GESTION.G_Fecha_Vta = Today.ToString("yyyyMMdd")
        GESTION.G_Hora_Vta = TimeOfDay.ToString("HHmmss")
        GESTION.G_Tiempo_PostView = lblHora.Text + ":" + lblMinutos.Text + ":" + lblSegundos.Text
        If perfil <> "Regrabador" Then

            GESTION.G_IdLlamada = WS_CALL_ID
            biCliente.GuardaDatosCliente(CLIENTE, GESTION)
            InsertaAdicionales()
            InsertaBeneficiarios()

            MsgBox("Fin de la gestión. Presione ACEPTAR para continuar con el siguiente registro.", MsgBoxStyle.Information, "CORDIALPHONE")
            limpiarPrimeraPantalla()
            Buscar_Cliente()
        Else
            RECUPERACION.R_Estado = "T"
            RECUPERACION.R_Venta = "1"
            RECUPERACION.R_Fecha_Vta = Today.ToString("yyyyMMdd")
            RECUPERACION.R_Hora_Vta = TimeOfDay.ToString("HHmmss")
            RECUPERACION.R_IdLlamada = WS_CALL_ID
            'CLIENTE.CLI_ESTADO_OBJECION_CALIDAD = 1
            biGesRes.GuardaClienteGes(CLIENTE, RECUPERACION, GESTION)
            biGesRes.ActualizaClienteVenta(CLIENTE, GESTION)

            InsertaAdicionales()
            InsertaBeneficiarios()

            MsgBox("Fin de la gestión. Presione ACEPTAR para salir del formulario.", MsgBoxStyle.Information, csNombreAplicacion)
            limpiarPrimeraPantalla()
            Me.Hide()
            frmRegrabaciones.ShowDialog()
            BuscaGes()
        End If

    End Sub

    Private Function LlenaParentescoCondicion() As Boolean
        LlenaParentescoCondicion = False
        Dim tmpParentesco As New List(Of eTipoParentesco)
        Dim listParentesco As New List(Of eTipoParentesco)
        Dim entParentesco As New eTipoParentesco

        Dim biPlan As New clsPlanBI
        entParentesco.idParentesco = 0
        entParentesco.nombreParentesco = "---No ingresado---"
        tmpParentesco.Add(entParentesco)
        tmpParentesco.Clear()

        entParentesco.idParentesco = cmbParentescoAdic.SelectedValue
        'Dim count As Int16 = 0
        If cmbParentescoAdic.SelectedIndex <> -1 And cmbParentescoAdic.SelectedIndex <> 0 Then
            listParentesco = biClsParentescoCampania.BuscaParentescoEdadPorId(vgCampania.idCRM, entParentesco.idParentesco)

            For y As SByte = 0 To listParentesco.Count - 1
                If entParentesco.idParentesco = listParentesco(y).idParentesco Then
                    entParentesco.edadMin = listParentesco(y).edadMin
                    entParentesco.edadMax = listParentesco(y).edadMax

                    planE = biClsPlan.BuscarPlanPorIdPlan(listParentesco(y).idPlan)

                    Dim count As Int16 = 0

                    Dim edadCliente As Int16 = edad(dtFechaNacAdic.Value) 'DateDiff(DateInterval.Year, Date.Parse(dtFechaNac.Text), Date.UtcNow)

                    Select Case edadCliente
                        Case entParentesco.edadMin To entParentesco.edadMax
                            count = 1
                            LlenaParentescoCondicion = True
                    End Select

                    If count = 1 Then
                        ufAdic = planE.primaUF
                        idPlanAdic = planE.idPlan
                        totalUfAdic = CDbl(totalUfAdic) + CDbl(ufAdic)
                        vgFuncionComun.sumaUFAdicionales()
                    End If
                    'Next
                End If
            Next
        Else
            MsgBox("Debe ingresar un parentezco.", vbExclamation, csNombreAplicacion)
        End If

        Return LlenaParentescoCondicion
    End Function

    Private Sub cmbMedioPago_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbMedioPago.SelectedIndexChanged
        Dim biMedioPago As New clsMedioPagoBI
        medioPagoGlobal = biMedioPago.BuscaDatosMedioPagoPorId(vgCampania.idCRM, CLIENTE.C_Id, cmbMedioPago.SelectedValue)

        If medioPagoGlobal.otroMedioPago Then
            Panelotro.Visible = True
            txtUltDigitos.Visible = False
            lblUltDig.Visible = False
            CmbBanco.SelectedIndex = 0
            cmbMes.SelectedIndex = 0
            cmbAnio.SelectedIndex = 0
            cmbTpoTarjeta.SelectedIndex = 0
            txtNumeroCta.Text = ""
        Else
            Panelotro.Visible = False
            txtUltDigitos.Text = medioPagoGlobal.NumeroMedioPago
            txtUltDigitos.Visible = True
            lblUltDig.Visible = True
        End If
    End Sub
    Private Sub ValidaNumero_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNroV.KeyPress, txtCelular.KeyPress, txtRutV.KeyPress, txtFonoVenta.KeyPress, txtCelular.KeyPress, txtUltDigitos.KeyPress, txtNumeroCta.KeyPress
        vgFuncionComun.validaNumeros(e)
    End Sub
    Private Sub txtCalle_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCalleV.KeyPress, txtNombreV.KeyPress, txtPaternoV.KeyPress, txtMaternoV.KeyPress, txtReferenciaV.KeyPress, txtEmail.KeyPress
        vgFuncionComun.validaCaracter(e)
    End Sub

    Private Sub txtRutcontact_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub btnExclusiones_Click(sender As Object, e As EventArgs) Handles btnExclusiones.Click
        frmExclusiones.ShowDialog()
    End Sub

    Private Sub ComboBoxReconoce_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbReconoce.SelectedIndexChanged
        If cmbReconoce.SelectedIndex = 1 Then
            CmbRealiza.Visible = True
            lblRealiza.Visible = True
        Else
            CmbRealiza.Visible = False
            lblRealiza.Visible = False
        End If
    End Sub

    Private Sub tmLabelRegrabacion_Tick(sender As Object, e As EventArgs) Handles tmLabelRegrabacion.Tick
        If lblRegrabacion.ForeColor = Color.Red Then
            lblRegrabacion.ForeColor = Color.Yellow
        Else
            lblRegrabacion.ForeColor = Color.Red
        End If
    End Sub

    Dim TiempoInicioPostView As DateTime


    Private Sub btnDescanso_Click(sender As Object, e As EventArgs) Handles btnDescanso.Click
        frmDescanso.ShowDialog()
    End Sub

    Private Sub TmrPostView_Tick(sender As Object, e As EventArgs) Handles tmrPostView.Tick

        lblSegundos.Text += 1
        If lblSegundos.Text.Length = 1 Then lblSegundos.Text = "0" + lblSegundos.Text
        If lblSegundos.Text = "60" Then
            lblMinutos.Text += 1
            If lblMinutos.Text.Length = 1 Then lblMinutos.Text = "0" + lblMinutos.Text
            lblSegundos.Text = 0
        End If
        If lblMinutos.Text = 60 Then
            lblHora.Text += 1
            lblMinutos.Text = 0
        End If

    End Sub

    Public Sub sumaUFAdicionales()
        Dim uf As Double = 0
        Dim pesos As Int64 = 0
        totalUfAdic = 0
        TotalPesos = 0
        For i As Integer = 0 To lstAdi.Count - 1
            uf = CDbl(lstAdi.Item(i).a_primaUf)
            pesos = CDbl(lstAdi.Item(i).a_primaPesos)
            totalUfAdic = totalUfAdic + uf
            TotalPesos = TotalPesos + pesos
        Next
        Dim plan As New ePlan
        plan = vgListPlanes.Find(Function(tmpC As ePlan) tmpC.idPlan = cmbPlan.SelectedValue)
        ePlanGlobal.PrimaPesos = plan.PrimaPesos
        ePlanGlobal.primaUF = plan.primaUF

        lblPrimaUF.Text = (totalUfAdic + CDbl(plan.primaUF)).ToString

        TotalPesos = CInt(TotalPesos + plan.PrimaPesos)
        lblPrimaPesos.Text = TotalPesos.ToString

    End Sub

    Private Sub CmbBanco_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbBanco.SelectedIndexChanged
        If CmbBanco.Text.Trim() = "BANCO DE CREDITO E INVERSIONES" Then
            cmbMes.Visible = False
            Label96.Visible = False
            Label99.Visible = False
            cmbAnio.Visible = False
            txtTpoTarjeta.Visible = False
            cmbTpoTarjeta.Visible = False
        Else
            cmbMes.Visible = True
            Label96.Visible = True
            Label99.Visible = True
            cmbAnio.Visible = True
            txtTpoTarjeta.Visible = True
            cmbTpoTarjeta.Visible = True
        End If
    End Sub
End Class