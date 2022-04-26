Imports System.Collections.Generic
Imports Entidad, BI
Public Class frmAdicionales

    Dim biClsTipoContrato As New clsTipoContratoBI
    Dim biClsParentescoCampania As New clsParentescoCampaniaBI
    Dim biClsPlan As New clsPlanBI

    Private Sub frmAdicionales_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try
            limpiaAdicionales()
            dtAdicional.DataSource = Nothing
            dtAdicional.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub frmAdicionales_Load(sender As Object, e As EventArgs) Handles Me.Load
        vgFuncionComun.llenaComboBox(cmbParentescoAdic, "nombreParentesco", "idParentesco", listParentescoCampania.ToArray)
        If lstAdi.Count > 0 Then
            cargaAdicionales()
            sumaUFAdicionales()
        End If
    End Sub

    Private Sub btnAgregar_Click(sender As Object, e As EventArgs) Handles btnAgregar.Click
        If ValAdicional(0) Then
            If LlenaParentescoCondicion() <> True Then
                Exit Sub
            End If
            If paInsertaAdicional() Then
            End If
        End If
        sumaUFAdicionales()
    End Sub

    Private Sub CmdSiguiente_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSiguiente.Click
        insertaAdicionalesGrilla()
        limpiaAdicionales()
        Me.Close()
    End Sub

    Private Function insertaAdicionalesGrilla() As Boolean
        lstAdi.Clear()

        For I As Int16 = 0 To dtAdicional.Rows.Count - 1
            Dim adi As New eAdicional
            adi.A_Nombre = IIf(dtAdicional.Rows(I).Cells("nombre").Value Is DBNull.Value, Nothing, dtAdicional.Rows(I).Cells("nombre").Value)
            adi.A_Paterno = IIf(dtAdicional.Rows(I).Cells("paterno").Value Is DBNull.Value, Nothing, dtAdicional.Rows(I).Cells("paterno").Value)
            adi.A_Materno = IIf(dtAdicional.Rows(I).Cells("materno").Value Is DBNull.Value, Nothing, dtAdicional.Rows(I).Cells("materno").Value)
            adi.A_Rut = IIf(dtAdicional.Rows(I).Cells("rut").Value Is DBNull.Value, Nothing, dtAdicional.Rows(I).Cells("rut").Value)
            adi.A_Dv = IIf(dtAdicional.Rows(I).Cells("dv").Value Is DBNull.Value, Nothing, dtAdicional.Rows(I).Cells("dv").Value)
            adi.A_TipoBeneficiario = IIf(dtAdicional.Rows(I).Cells("idParentescoAdi").Value Is DBNull.Value, Nothing, dtAdicional.Rows(I).Cells("idParentescoAdi").Value)
            adi.A_NombreBeneficiario = IIf(dtAdicional.Rows(I).Cells("tipo_parentesco").Value Is DBNull.Value, Nothing, dtAdicional.Rows(I).Cells("tipo_parentesco").Value)
            adi.A_Sexo = IIf(dtAdicional.Rows(I).Cells("sexo").Value Is DBNull.Value, Nothing, dtAdicional.Rows(I).Cells("sexo").Value)
            adi.A_FechaNacimiento = IIf(dtAdicional.Rows(I).Cells("fechaNacimiento").Value Is DBNull.Value, Nothing, dtAdicional.Rows(I).Cells("fechaNacimiento").Value)
            adi.a_primaUf = IIf(dtAdicional.Rows(I).Cells("ValorUf").Value Is DBNull.Value, Nothing, dtAdicional.Rows(I).Cells("ValorUf").Value)
            adi.a_primaPesos = IIf(dtAdicional.Rows(I).Cells("valorPesos").Value Is DBNull.Value, Nothing, dtAdicional.Rows(I).Cells("valorPesos").Value)
            adi.idPlanAdic = IIf(dtAdicional.Rows(I).Cells("idPlan").Value Is DBNull.Value, Nothing, dtAdicional.Rows(I).Cells("idPlan").Value)
            adi.a_salud = IIf(dtAdicional.Rows(I).Cells("SistemaSalud").Value Is DBNull.Value, Nothing, dtAdicional.Rows(I).Cells("SistemaSalud").Value)

            lstAdi.Add(adi)
        Next I
        Return True
    End Function

    Private Sub paModificaAdicional()
        LlenaParentescoCondicion()
        If dtAdicional.Rows.Count <> 0 Then
            dtAdicional.Rows(numeroFila).Cells("nombre").Value = txtNombreA.Text
            dtAdicional.Rows(numeroFila).Cells("paterno").Value = txtPaternoA.Text
            dtAdicional.Rows(numeroFila).Cells("materno").Value = txtMaternoA.Text
            dtAdicional.Rows(numeroFila).Cells("rut").Value = txtRutA.Text
            dtAdicional.Rows(numeroFila).Cells("sexo").Value = cmbSexoA.Text
            dtAdicional.Rows(numeroFila).Cells("dv").Value = txtDvA.Text
            dtAdicional.Rows(numeroFila).Cells("tipo_parentesco").Value = cmbParentescoAdic.Text
            dtAdicional.Rows(numeroFila).Cells("idParentescoAdi").Value = cmbParentescoAdic.SelectedValue
            dtAdicional.Rows(numeroFila).Cells("ValorUf").Value = ufAdic
            dtAdicional.Rows(numeroFila).Cells("valorPesos").Value = PesosAdic
            dtAdicional.Rows(numeroFila).Cells("SistemaSalud").Value = cmbSaludAdic.Text
            dtAdicional.Rows(numeroFila).Cells("fechaNacimiento").Value = dtFechaNacAdic.Value.ToString("yyyy-MM-dd")
        Else
            MsgBox("Debe agregar Primero antes de Modificar!.", vbInformation, csNombreAplicacion)
            Exit Sub
        End If

        limpiaAdicionales()
        btnModificar.Enabled = True
        btnAgregar.Enabled = True
    End Sub

    Private Function LlenaParentescoCondicion() As Boolean
        LlenaParentescoCondicion = False
        Dim listParentesco As New List(Of eTipoParentesco)
        Dim entParentesco As New eTipoParentesco

        Dim biPlan As New clsPlanBI
        entParentesco.idParentesco = 0
        entParentesco.nombreParentesco = "---No ingresado---"

        entParentesco.idParentesco = cmbParentescoAdic.SelectedValue
        If cmbParentescoAdic.SelectedIndex <> -1 And cmbParentescoAdic.SelectedIndex <> 0 Then
            listParentesco = biClsParentescoCampania.BuscaParentescoEdadPorId(vgCampania.idCRM, entParentesco.idParentesco)

            For y As SByte = 0 To listParentesco.Count - 1
                If entParentesco.idParentesco = listParentesco(y).idParentesco Then
                    entParentesco.edadMin = listParentesco(y).edadMin
                    entParentesco.edadMax = listParentesco(y).edadMax

                    planE = biClsPlan.BuscarPlanPorIdPlan(listParentesco(y).idPlan)

                    Dim edadCliente As Int16 = edad(dtFechaNacAdic.Value)

                    Select Case edadCliente
                        Case entParentesco.edadMin To entParentesco.edadMax
                            ufAdic = planE.primaUF
                            PesosAdic = planE.PrimaPesos
                            idPlanAdic = planE.idPlan
                            totalUfAdic = CDbl(totalUfAdic) + CDbl(ufAdic)
                            valorPesosUf = planE.PrimaPesos
                            Return True
                    End Select

                End If
            Next
        Else
            MsgBox("Debe ingresar un parentezco.", vbExclamation, csNombreAplicacion)
        End If

        Return LlenaParentescoCondicion
    End Function

    Private Function ValAdicional(ByVal num As Integer) As Boolean

        If txtNombreA.Text = "" Then
            MsgBox("El campo nombre es obligatorio.", vbExclamation, csNombreAplicacion)
            txtNombreA.Focus()
            Exit Function
        ElseIf Len(txtNombreA.Text) < 2 Then
            MsgBox("El nombre ingresado No es valido.", vbExclamation, csNombreAplicacion)
            txtNombreA.Focus()
            Exit Function
        End If

        If txtPaternoA.Text = "" Then
            MsgBox("El campo Apellido Paterno es obligatorio.", vbExclamation, csNombreAplicacion)
            txtPaternoA.Focus()
            Exit Function
        ElseIf Len(txtPaternoA.Text) < 2 Then
            MsgBox("El Apellido Paterno ingresado, No es valido.", vbExclamation, csNombreAplicacion)
            txtPaternoA.Focus()
            Exit Function
        End If
        If txtMaternoA.Text = "" Then
            MsgBox("El campo Apellido Materno es obligatorio.", vbExclamation, csNombreAplicacion)
            txtMaternoA.Focus()
            Exit Function
        ElseIf Len(txtMaternoA.Text) < 2 Then
            MsgBox("El Apellido Materno ingresado, No es valido.", vbExclamation, csNombreAplicacion)
            txtMaternoA.Focus()
            Exit Function
        End If

        If cmbSexoA.SelectedIndex = -1 Or cmbSexoA.SelectedIndex = 0 Then
            MsgBox("Debe seleccionar el sexo.", vbExclamation, csNombreAplicacion)
            cmbSexoA.Focus()
            Exit Function
        End If

        If cmbParentescoAdic.SelectedIndex = -1 Or cmbParentescoAdic.SelectedIndex = 0 Then
            MsgBox("Debe seleccionar el parentesco de la carga.", vbExclamation, csNombreAplicacion)
            cmbParentescoAdic.Focus()
            Exit Function
        End If

        If cmbSaludAdic.SelectedIndex < 1 Then
            MsgBox("Debe seleccionar el sistema de salud del adicional.", vbExclamation, csNombreAplicacion)
            cmbSaludAdic.Focus()
            Exit Function
        End If

        If txtRutA.Text <> "" Then
            '    MsgBox("Debe ingresar rut del Adicional", MsgBoxStyle.Exclamation, csNombreAplicacion)
            '    txtRutA.Focus()
            '    Exit Function
            'Else
            If Not vgFuncionComun.validarRut(Trim(txtRutA.Text) & "-" & Trim(txtDvA.Text)) Then
                MsgBox("El RUT del Adicional no es valido.", vbInformation, "CORDIALPHONE.")
                txtRutA.Focus()
                Exit Function
            End If
        End If

        If txtRutA.Text = GESTION.G_Rut Then
            MsgBox("El RUT del Adicional no puede ser igual al rut del titular.", vbInformation, "CORDIALPHONE.")
            txtRutA.Focus()
            Exit Function
        End If

        If num = 0 Then
            For x As Int16 = 0 To dtAdicional.Rows.Count - 1
                If dtAdicional.Rows(x).Cells("rut").Value = txtRutA.Text Then
                    MsgBox("Titular ya se encuentra ingresado con ese rut", vbInformation, "CORDIALPHONE")
                    txtRutA.Focus()
                    Exit Function
                End If
            Next
        End If

        ValAdicional = True
    End Function

    Private Function paInsertaAdicional() As Boolean
        FilaAgregar = dtAdicional.Rows.Count
        dtAdicional.Rows.Add(1)
        dtAdicional.Item("nombre", FilaAgregar).Value = txtNombreA.Text
        dtAdicional.Item("paterno", FilaAgregar).Value = txtPaternoA.Text
        dtAdicional.Item("materno", FilaAgregar).Value = txtMaternoA.Text
        dtAdicional.Item("rut", FilaAgregar).Value = txtRutA.Text
        dtAdicional.Item("dv", FilaAgregar).Value = txtDvA.Text
        dtAdicional.Item("Sexo", FilaAgregar).Value = cmbSexoA.Text
        dtAdicional.Item("idParentescoAdi", FilaAgregar).Value = cmbParentescoAdic.SelectedValue
        dtAdicional.Item("tipo_parentesco", FilaAgregar).Value = cmbParentescoAdic.Text
        dtAdicional.Item("fechaNacimiento", FilaAgregar).Value = dtFechaNacAdic.Value.ToString("yyyy-MM-dd")
        dtAdicional.Item("valorUf", FilaAgregar).Value = ufAdic.ToString
        dtAdicional.Item("valorPesos", FilaAgregar).Value = PesosAdic.ToString
        dtAdicional.Item("idPlan", FilaAgregar).Value = idPlanAdic.ToString
        dtAdicional.Item("SistemaSalud", FilaAgregar).Value = cmbSaludAdic.Text.ToString
        limpiaAdicionales()
        paInsertaAdicional = True
    End Function

    Private Sub limpiaAdicionales()

        txtNombreA.Text = ""
        txtPaternoA.Text = ""
        txtMaternoA.Text = ""
        txtRutA.Text = ""
        txtDvA.Text = ""
        dtFechaNacAdic.Value = dtFechaNacAdic.MinDate
        cmbParentescoAdic.SelectedIndex = 0
        cmbSexoA.SelectedIndex = 0
        cmbSaludAdic.SelectedIndex = 0
        lblPesosUF.Text = ""
        lblTotalPrimaUf.Text = ""
    End Sub

    Private Sub cargaAdicionales()
        dtAdicional.DataSource = Nothing
        dtAdicional.Rows.Clear()
        For i As Int16 = 0 To lstAdi.Count - 1
            dtAdicional.Rows.Add()
            dtAdicional.Rows(i).Cells("nombre").Value = lstAdi.Item(i).A_Nombre
            dtAdicional.Rows(i).Cells("paterno").Value = lstAdi.Item(i).A_Paterno
            dtAdicional.Rows(i).Cells("materno").Value = lstAdi.Item(i).A_Materno
            dtAdicional.Rows(i).Cells("rut").Value = lstAdi.Item(i).A_Rut
            dtAdicional.Rows(i).Cells("dv").Value = lstAdi.Item(i).A_Dv
            dtAdicional.Rows(i).Cells("idParentescoAdi").Value = lstAdi.Item(i).A_TipoBeneficiario
            dtAdicional.Rows(i).Cells("tipo_parentesco").Value = lstAdi.Item(i).A_NombreBeneficiario
            dtAdicional.Rows(i).Cells("sexo").Value = lstAdi.Item(i).A_Sexo
            dtAdicional.Rows(i).Cells("fechaNacimiento").Value = lstAdi.Item(i).A_FechaNacimiento
            dtAdicional.Rows(i).Cells("ValorUf").Value = lstAdi.Item(i).a_primaUf
            dtAdicional.Rows(i).Cells("ValorPesos").Value = lstAdi.Item(i).a_primaPesos
            dtAdicional.Rows(i).Cells("idPlan").Value = lstAdi.Item(i).idPlanAdic
            dtAdicional.Rows(i).Cells("SistemaSalud").Value = lstAdi.Item(i).a_salud
        Next
    End Sub

    Private Sub dtAdicional_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dtAdicional.CellClick
        If e.RowIndex = -1 Then Exit Sub

        txtNombreA.Text = dtAdicional.Rows(e.RowIndex).Cells("nombre").Value
        txtPaternoA.Text = dtAdicional.Rows(e.RowIndex).Cells("paterno").Value
        txtMaternoA.Text = dtAdicional.Rows(e.RowIndex).Cells("materno").Value
        txtRutA.Text = dtAdicional.Rows(e.RowIndex).Cells("rut").Value
        txtDvA.Text = dtAdicional.Rows(e.RowIndex).Cells("dv").Value
        cmbParentescoAdic.Text = dtAdicional.Rows(e.RowIndex).Cells("idParentescoAdi").Value
        cmbParentescoAdic.Text = dtAdicional.Rows(e.RowIndex).Cells("tipo_parentesco").Value
        dtFechaNacAdic.Value = dtAdicional.Rows(e.RowIndex).Cells("fechaNacimiento").Value
        cmbSexoA.Text = dtAdicional.Rows(e.RowIndex).Cells("Sexo").Value
        cmbSaludAdic.Text = dtAdicional.Rows(e.RowIndex).Cells("SistemaSalud").Value

        numeroFila = e.RowIndex
    End Sub

    Private Sub btnModificar_Click(sender As Object, e As EventArgs) Handles btnModificar.Click
        'If LlenaParentescoCondicion() <> True Then
        '    Exit Sub
        'End If
        If ValAdicional(1) Then
            LlenaParentescoCondicion()
            paModificaAdicional()
        End If
    End Sub

    Private Function paEliminaAdicional() As Boolean
        paEliminaAdicional = False
        For i As Int16 = 0 To dtAdicional.Rows.Count - 1
            If dtAdicional.Rows(i).Cells("rut").Value = txtRutA.Text Then
                dtAdicional.Rows.RemoveAt(i)
                paEliminaAdicional = True
                vgFuncionComun.sumaUFAdicionales()
                Exit Function
            End If
        Next
    End Function

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click
        If (MsgBox("Desea eliminar el adicional seleccionado", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation, csNombreAplicacion)) = MsgBoxResult.No Then Exit Sub

        If paEliminaAdicional() Then
            MsgBox("Adicional eliminado", MsgBoxStyle.Information, csNombreAplicacion)
        Else
            MsgBox("No se ha podido eliminar el adicional, favor verificar rut", MsgBoxStyle.Exclamation, csNombreAplicacion)
            txtRutA.Focus()
        End If
        sumaUFAdicionales()
    End Sub

    Public Sub sumaUFAdicionales()
        Dim uf As Double = 0
        Dim pesos As Int64 = 0
        totalUfAdic = 0
        TotalPesos = 0
        For i As Integer = 0 To dtAdicional.RowCount - 1
            uf = CDbl(dtAdicional.Item("valorUf", i).Value)
            pesos = CDbl(dtAdicional.Item("valorPesos", i).Value)
            totalUfAdic = totalUfAdic + uf
            TotalPesos = TotalPesos + pesos
        Next

        lblTotalPrimaUf.Text = "Total UF: " + (totalUfAdic + CDbl(ePlanGlobal.primaUF)).ToString
        TotalPesos = CInt(TotalPesos + ePlanGlobal.PrimaPesos)
        lblPesosUF.Text = "Valor Pesos: " + TotalPesos.ToString

    End Sub
End Class