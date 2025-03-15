'en esta clase dse manejaran las actualizaciones de la tablas repportes de excel relacionadas a las entidades de autocad 
'su manejo sera puramente de excel y las demas librerias deberia cargase en clases correspondinetes para su manejo 
'Primera funccion obtendra las celda activa del libro activo GetACRNG() as range
'segunda funcion determinara si esta celda pertenece a una tabla IsRNGInTbl(RNG as RAnge) as boolean 
Imports Microsoft.Office.Interop

Public Class TblManagementExc
    Private excelApp As Excel.Application

    ' Constructor para inicializar la aplicación de Excel
    Public Sub New()
        Try
            ' Crea una nueva instancia de Excel.Application
            ' Opcional: Hacer visible Excel (útil para depuración)
            excelApp = New Excel.Application With {
                .Visible = True
            }
        Catch ex As Exception
            MsgBox("Error al inicializar Excel: " & ex.Message)
        End Try
    End Sub

    ' Función para obtener la celda activa del libro activo en Excel
    Public Function GetACRNG() As Excel.Range
        Try
            ' Verifica si hay un libro activo en Excel
            If excelApp.ActiveWorkbook Is Nothing Then
                Throw New Exception("No hay un libro activo en Excel.")
            End If

            ' Verifica si hay una hoja activa
            If excelApp.ActiveSheet Is Nothing Then
                Throw New Exception("No hay una hoja activa en Excel.")
            End If

            ' Obtiene la celda activa
            Return excelApp.ActiveCell
        Catch ex As Exception
            ' Manejo de excepciones
            MsgBox("Error al obtener la celda activa: " & ex.Message)
            Return Nothing
        End Try
    End Function

    ' Función para determinar si una celda pertenece a una tabla
    Public Function IsRNGInTbl(RNG As Excel.Range) As Boolean
        Try
            ' Verifica si el rango es válido
            If RNG Is Nothing Then
                Throw New Exception("El rango proporcionado es nulo.")
            End If

            ' Verifica si la celda está en una tabla
            If RNG.ListObject IsNot Nothing Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            ' Manejo de excepciones
            MsgBox("Error al verificar si el rango está en una tabla: " & ex.Message)
            Return False
        End Try
    End Function

    ' Destructor para cerrar la aplicación de Excel cuando el objeto es eliminado
    Protected Overrides Sub Finalize()
        Try
            If excelApp IsNot Nothing Then
                excelApp.Quit()
                excelApp = Nothing
            End If
        Catch ex As Exception
            MsgBox("Error al cerrar Excel: " & ex.Message)
        Finally
            MyBase.Finalize()
        End Try
    End Sub
End Class

