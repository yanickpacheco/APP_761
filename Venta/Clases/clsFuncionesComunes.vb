Imports System.Net
Imports System.Data.SqlClient
Imports System.Reflection
Imports System.Collections.Generic
Imports Entidad

Public Class clsFuncionesComunes

    Public Sub llenaComboBox(ByRef _drd As ComboBox, ByRef _strDisplay As String, ByRef _strValue As String, ByRef _origenDato As Array)
        _drd.DisplayMember = _strDisplay
        _drd.ValueMember = _strValue
        _drd.DataSource = _origenDato
    End Sub

    Public Sub llenaComboboxTabla(ByVal cmbParCombo As ComboBox, ByVal dt As DataTable, ByVal strDisplay As String, ByVal strValue As String)

        With cmbParCombo
            .ValueMember = strValue
            .DisplayMember = strDisplay
            .DataSource = dt
        End With

    End Sub

    Public Sub llenaCheckBox(ByRef _drd As CheckedListBox, ByRef _origenDato As List(Of eExclusion))

        For i As Integer = 0 To _origenDato.Count - 1
            _drd.Items.Add(_origenDato.Item(i).Exclusion)
        Next
    End Sub

    Public Sub llenaGrilla(ByRef _grd As DataGridView, ByRef _origenDato As Array)
        _grd.DataSource = _origenDato
    End Sub

    Public Function getIp() As String

        Dim ip As System.Net.IPHostEntry
        ip = Dns.GetHostEntry(My.Computer.Name)

        Dim valorIp As String

        valorIp = ""

        For Each valor As IPAddress In ip.AddressList
            valorIp = valor.ToString()
        Next

        getIp = valorIp

    End Function
#Region "Validacion de Edad"
    ''' <summary>
    ''' Calculo de edad exacta
    ''' </summary>
    ''' <param name="fecNacimiento"></param>
    ''' <returns>Devuelve edad segun dia, mes y año</returns>
    ''' <remarks></remarks>
    ''' 

    Public Function edad(ByVal fecNacimiento As Date) As Integer
        Dim vEdad As Integer
        vEdad = Today.Year - fecNacimiento.Year
        If Today.Month < fecNacimiento.Month Then
            vEdad = vEdad - 1
        ElseIf Today.Month = fecNacimiento.Month And Today.Day < fecNacimiento.Day Then
            vEdad = vEdad - 1
        End If
        Return vEdad
    End Function
    ''' <summary>
    ''' Establece si la Fec Nacimiento es mayor o igual a la Edad Minima
    ''' </summary>
    ''' <param name="fecNacimiento"></param>
    ''' <param name="EdadMin"></param>
    ''' <returns>Devuelve verdadero si la Fecha es mayor o igual a la edad minima </returns>
    ''' <remarks></remarks>
    Public Function edad(ByVal fecNacimiento As Date, ByVal EdadMin As Int16) As Boolean
        Dim vEdad As Integer
        vEdad = edad(fecNacimiento)
        Return vEdad >= EdadMin
    End Function
    ''' <summary>
    ''' Establece si la Fec Nacimiento esta entre la edad Minima y maxima
    ''' </summary>
    ''' <param name="fecNacimiento"></param>
    ''' <param name="EdadMin"></param>
    ''' <param name="EdadMax"></param>
    ''' <returns>Devuelve verdadero si la Fecha esta entre la edad minima y edad maxima</returns>
    ''' <remarks></remarks>
    Public Function edad(ByVal fecNacimiento As Date, ByVal EdadMin As Int16, ByVal EdadMax As Int16) As Boolean
        Dim vEdad As Integer
        vEdad = edad(fecNacimiento)
        Return (vEdad >= EdadMin And vEdad <= EdadMax)
    End Function
#End Region


    Public Function ValidaEmail(ByVal email As String) As Boolean
        Dim emailRegex As New System.Text.RegularExpressions.Regex("^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$")
        Dim emailMatch As System.Text.RegularExpressions.Match = emailRegex.Match(email)
        Return emailMatch.Success
    End Function

    Public Function proceParametroSendQuery(ByVal strQuery As String) As DataSet

        Dim cnn As New SqlConnection("data source=" & vgCampania.codigo & "; initial catalog=BDD_4942;User Id=aspnet;Password=123;")
        Dim ds As New DataSet
        Try

            Dim myCommand As New SqlDataAdapter(strQuery, cnn)
            myCommand.Fill(ds)

            Return ds

        Catch ex As Exception
            Return ds
        Finally
            cnn.Close()
        End Try

    End Function

    Public Sub LlenaCombos(ByRef cmb As ComboBox, ByVal dt As DataTable, ByVal strDisplay As String, ByVal strValue As String)

        With cmb
            .DataSource = dt
            .DisplayMember = strDisplay
            .ValueMember = strValue
        End With
    End Sub

    Public Sub llenaListBox(ByVal lboxListBox As ListBox, ByVal dt As DataTable, ByVal strDisplay As String, ByVal strValue As String)

        With lboxListBox
            .ValueMember = strValue
            .DisplayMember = strDisplay
            .DataSource = dt
        End With

    End Sub

    Public Sub llenaTablas(ByVal dt As DataTable, ByRef tabla As DataGridView)

        tabla.DataSource = Nothing
        tabla.DataSource = dt

    End Sub

    Public Function getDataTableDistinct(ByVal _tmpDT As DataTable, ByVal _camposDistinct As String()) As DataTable

        Try
            Return _tmpDT.DefaultView.ToTable(True, _camposDistinct)
        Catch ex As Exception
            Return New DataTable
        End Try

    End Function

    Public Sub sumaUFAdicionales()
        Dim uf As Double = 0
        Dim pesos As Int64 = 0
        totalUfAdic = 0
        TotalPesos = 0
        For i As Integer = 0 To frmAdicionales.dtAdicional.RowCount - 1
            uf = CDbl(frmAdicionales.dtAdicional.Item("valorUf", i).Value)
            pesos = CDbl(frmAdicionales.dtAdicional.Item("valorPesos", i).Value)
            totalUfAdic = totalUfAdic + uf
            TotalPesos = TotalPesos + pesos
        Next

        frmVenta.lblTotalPrimaUf.Text = "Total UF: " + (totalUfAdic + CDbl(ePlanGlobal.primaUF)).ToString

        TotalPesos = CInt(TotalPesos + ePlanGlobal.PrimaPesos)
        frmVenta.lblPesosUF.Text = "Valor Pesos: " + TotalPesos.ToString

    End Sub

    Public Function validaNumeros(ByVal e As System.Windows.Forms.KeyPressEventArgs) As System.Windows.Forms.KeyPressEventArgs

        If Char.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsSeparator(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If

        Return e

    End Function

    Public Function validaCaracter(ByVal e As System.Windows.Forms.KeyPressEventArgs) As System.Windows.Forms.KeyPressEventArgs

        If Char.IsLetterOrDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar <> "'" And e.KeyChar <> "*" Then
            e.Handled = False
        Else
            e.Handled = True
        End If

        Return e

    End Function


    ''' <summary>
    ''' Realiza la limpieza de los controles de los formularios de beneficiario como adicionales
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub LimpiaFormilarioBenAdi()

        'lstAdi.Clear()
        'dtAdicional.Rows.Clear()
        lstBen.Clear()
        frmBeneficiario.dtBeneficiario.Rows.Clear()

        If (perfil = "Regrabador") Then
            'dtAdicional.DataSource = Nothing
            frmBeneficiario.dtBeneficiario.DataSource = Nothing
        End If

    End Sub

    Public Function validarRut(ByRef rut As String) As Boolean
        Dim numRut, auxRut As Object
        Dim acumula As Double
        Dim correlativo, largoR, genera As Object
        Dim VER As Short
        Dim digVal As String

        largoR = InStr(1, Trim(rut), "-", CompareMethod.Binary) - 1
        numRut = Mid(Trim(rut), 1, largoR)
        digVal = Mid(Trim(rut), InStr(1, Trim(rut), "-", CompareMethod.Binary) + 1)
        auxRut = numRut
        acumula = 0
        genera = 2
        For correlativo = largoR To 1 Step -1
            acumula = (Val(Mid(auxRut, correlativo, 1)) * genera) + acumula
            If genera = 7 Then
                genera = 2
            Else
                genera = genera + 1
            End If
        Next correlativo
        VER = 11 - (acumula Mod 11)
        If VER = 10 Then
            If digVal = "k" Or digVal = "K" Then
                validarRut = True
            Else
                validarRut = False
            End If
        ElseIf VER = 11 Then
            If digVal = "0" Then
                validarRut = True
            Else
                validarRut = False
            End If
        ElseIf VER < 10 Then
            If digVal = Trim(Str(VER)) Then
                validarRut = True
            Else
                validarRut = False
            End If
        End If
    End Function


End Class
