'crear una funcion que verifique si los hndle de la columna hndle de la tabla "CunetasGeneral" existen en el documento de autocad 
'en el archivo ""\\MSI-CARLOS\1.0-Mediciones\Resumen Hormigones Typsa Med-13-R3 03-09-2024 1420 CF.xlsm""
'F1(cACAdHelpers as ACAdHelpers)
'F2(TableName as String) as List(of ) esta devolvera el rango completo de la tabla en un listado de valores tipo (string, double, date) dependiendo de los valores encontrados y los formatos en las celdas
'MAintestCMD comando de autocad para poner a prueba el codigo commadMethod
'estos comentarios son descriptivos y deberán permanecer tal cual para la comprensión posterior del código
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.DatabaseServices
Imports System.Linq

' Clase encargada de verificar la existencia de hndle en AutoCAD

' Clase encargada de manejar funciones relacionadas con AutoCAD
Public Class AutoCADHelper
    ' F1: Verificar si los hndle existen en el documento de AutoCAD
    ' Devuelve dos listas: una de NoEncontrados y otra de Encontrados

    Public Shared Function VerificarHandles(hndleList As List(Of String)) As (List(Of String), List(Of String))
        Dim handlesEncontrados As New List(Of String)
        Dim handlesNoEncontrados As New List(Of String)

        ' Obtener el documento activo de AutoCAD
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            For Each handleStr As String In hndleList
                Try
                    ' Crear un objeto TabName a partir de la cadena y convertir a Hexadecimal
                    Dim handleObj As New Handle(Convert.ToInt64(handleStr, 16))
                    Dim objId As ObjectId = db.GetObjectId(False, handleObj, 0)

                    ' Verificar si el ObjectId es válido
                    If objId.IsValid Then
                        handlesEncontrados.Add(handleStr)
                    Else
                        handlesNoEncontrados.Add(handleStr)
                    End If
                Catch ex As Exception
                    handlesNoEncontrados.Add(handleStr)
                End Try
            Next
            trans.Commit()
        End Using

        ' Devolver ambas listas: Encontrados y No Encontrados
        Return (handlesEncontrados, handlesNoEncontrados)
    End Function
    ' Nueva función que recibe el docName y maneja todo el proceso
    Public Shared Function ProcesarHandlesDesdeDoc(docName As String, tableName As String) As (List(Of String), List(Of String))
        ' BUSCAR INFORMACIÓN DEL ACCESO EN EL NOMBRE DEL ARCHIVO 
        Dim RefAcceso As String = CodigoExtractor.ExtraerCodigo(docName, "ACC-\d+")

        ' Si no se encuentra el código, salir
        If String.IsNullOrEmpty(RefAcceso) Then
            Console.WriteLine("El archivo no presenta referencia de accesos.")
            Return (Nothing, Nothing)
        End If

        Dim ACCESO As Double = CDbl(Mid(RefAcceso, 5))

        ' Verificar si se ha encontrado un código de acceso válido
        If ACCESO = 0 Then
            Throw New Exception("No se pudo extraer el código de acceso del nombre del archivo.")
        End If

        ' Obtener los datos de hndle desde el archivo Excel
        Dim hndleData As List(Of Object) = ExcelHelper.ObtenerDatosTabla(tableName, ACCESO)

        ' Convertir la lista de Object a String
        Dim hndleList As New List(Of String) '= hndleData.Select(Function(obj) obj.ToString()).ToList()
        For Each obj As Object In hndleData
            hndleList.Add(obj.ToString())
        Next
        ' Verificar si los hndle existen en AutoCAD
        Return VerificarHandles(hndleList)
    End Function
End Class

' Clase principal para manejar el comando de AutoCAD
Public Class MainAutoCADCommand
    ' MAintestCMD: Comando para AutoCAD que prueba la funcionalidad
    <Autodesk.AutoCAD.Runtime.CommandMethod("MAintestCMD")>
    Public Shared Sub TestCommand()
        ' Get the document manager
        Dim docMgr As DocumentCollection = Application.DocumentManager
        Dim acDoc As Document = docMgr.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor

        ' Verificar si los hndle existen en AutoCAD
        Dim Listhandles As (List(Of String), List(Of String)) = AutoCADHelper.ProcesarHandlesDesdeDoc(acDoc.Name, "CunetasGeneral")

        ' Mostrar resultados en la ventana de AutoCAD
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor

        If Listhandles.Item1.Count = 0 Then
            ed.WriteMessage("Todos los hndle fueron encontrados.")
        Else
            ed.WriteMessage("Handles no encontrados: " & String.Join(", ", Listhandles.Item1))
        End If
    End Sub
End Class
