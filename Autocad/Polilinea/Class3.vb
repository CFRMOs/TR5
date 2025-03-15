' En esta clase se creara un proceso para chequear los handles listados en la columnas "TabName" de la tabla "CunetasGeneral"
'
Imports Microsoft.Office.Interop.Excel

Public Class Class3
    Sub ChAllandle()
        Dim Headers As New List(Of String) From {"TabName"}
        'Dim Data As List(Of Object) = ExRNGSe.SelectRowOnTbl("CunetasGeneral", Headers)
        Dim ExcelRS As New ExcelRangeSelector
        For Each Rng As Range In ExcelRS.GetColumnsRange()
            For Each c As Range In Rng.Cells
                If CLHandle.CheckIfExistHd(c.Value.ToString()) Then

                End If
            Next
        Next
    End Sub

End Class
