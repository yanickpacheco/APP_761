

Public Class clsFuncionesDatatable

    'Public Shared Function getDatatableSelect(ByVal _tmpDT As DataTable, ByVal _tmpCondicion As String, ByVal _campos As String) As DataTable

    '    Try
    '        Dim foundRows() As DataRow
    '        foundRows = _tmpDT.Select(_tmpCondicion)

    '        Return foundRows.CopyToDataTable


    '        'Return _tmpDT.Select(_tmpCondicion, _campos) ' .CopyToDataTable
    '    Catch ex As Exception
    '        Return New DataTable
    '    End Try

    'End Function

    'Public Shared Function getDatatableSelect(ByVal _tmpDT As DataTable, ByVal _tmpCondicion As String, ByVal _columnNames As String()) As DataTable

    '    Try
    '        Return _tmpDT.Select(_tmpCondicion).CopyToDataTable.DefaultView.ToTable(True, _columnNames)
    '    Catch ex As Exception
    '        Return New DataTable
    '    End Try

    'End Function
    'Public Shared Function getDatatableSelect(ByVal _tmpDT As DataTable, ByVal _tmpCondicion As String) As DataTable

    '    Try
    '        Return _tmpDT.Select(_tmpCondicion).CopyToDataTable '.DefaultView.ToTable(True, _columnNames)
    '    Catch ex As Exception
    '        Return New DataTable
    '    End Try

    'End Function
    'Public Shared Function getDatatableSelect(ByVal _tmpDT As DataTable, ByVal _columnNames As String()) As DataTable

    '    Try
    '        Return _tmpDT.DefaultView.ToTable(True, _columnNames)
    '    Catch ex As Exception
    '        Return New DataTable
    '    End Try

    'End Function

End Class
