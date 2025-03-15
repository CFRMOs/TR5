Imports System.Windows.Forms
Public Class GetControlsDGView
    Public Shared Function GetDGView(TabCtrol As TabControl) As DataGridView
        ' Verifica si el TabControl tiene una página seleccionada
        If TabCtrol.SelectedTab IsNot Nothing Then
            ' Accede a la página activa
            Dim activeTabPage As TabPage = TabCtrol.SelectedTab

            ' Busca el DataGridView dentro de la página activa
            For Each ctrl As Control In activeTabPage.Controls
                If TypeOf ctrl Is DataGridView Then
                    Return CType(ctrl, DataGridView)
                    Exit For
                End If
            Next
        Else
            'MessageBox.Show("No hay ninguna página activa seleccionada en el TabControl")
        End If
        Return Nothing
    End Function
    Public Shared Function GetDGViewByTabName(TabControl2 As TabControl, TabName As String) As DataGridView
        ' Verifica si el TabControl tiene una página seleccionada
        If TabControl2 IsNot Nothing Then
            ' Accede a la página activa
            For Each Tab As TabPage In TabControl2.TabPages
                If Tab.Text = TabName Then
                    ' Busca el DataGridView dentro de la página activa
                    For Each ctrl As Control In Tab.Controls
                        If TypeOf ctrl Is DataGridView Then
                            Return CType(ctrl, DataGridView)
                        End If
                    Next
                End If
            Next
        Else
            MessageBox.Show("No hay ninguna página activa seleccionada en el TabControl")
        End If
        Return Nothing
    End Function
    Public Shared Sub Delete(TabControl2 As TabControl, CAcadHelp As ACAdHelpers)
        Dim DGView As DataGridView = GetControlsDGView.GetDGView(TabControl2) 'TabCtrol)

        Dim selectedCells As DataGridViewSelectedCellCollection = DGView.SelectedCells
        ' Mostrar mensaje de verificación
        Dim result As DialogResult = MessageBox.Show("¿Estás seguro de que deseas eliminar esta entidad?", "Confirmación de eliminación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        For Each cell As DataGridViewCell In selectedCells
            Dim hdString = DGView.Rows(cell.RowIndex).Cells("Handle").Value.ToString()
            If hdString <> vbNullString Then
                If result = DialogResult.Yes Then
                    CLHandle.GetEntityByStrHandle(hdString)
                    If CLHandle.CheckIfExistHd(hdString) Then
                        CLHandle.EraseObjectByHandle(CLHandle.Chandle(hdString))
                        'DGView.Rows(cell.RowIndex)
                    End If
                End If
            End If
        Next
        UpdateDataDGV.CUPDateDGView(DGView, CAcadHelp)
    End Sub
End Class
