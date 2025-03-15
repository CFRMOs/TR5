Imports System.IO
Imports System.Diagnostics
Imports System.Windows.Forms

Public Class FileSearcher
    Private _carpetaBase As String

    ' Constructor
    Public Sub New(carpetaBase As String)
        _carpetaBase = carpetaBase
    End Sub

    ' Método para buscar un archivo en la carpeta y subcarpetas
    ' Se permite buscar cualquier tipo de archivo mediante la extensión especificada
    Public Function BuscarArchivo(nombreArchivo As String, Optional ext As String = "pdf", Optional profundidad As SearchOption = SearchOption.TopDirectoryOnly) As String
        Dim archivos As String() = Directory.GetFiles(_carpetaBase, "*" & nombreArchivo & "*." & ext, profundidad)
        Return If(archivos.Length > 0, archivos(0), String.Empty)
    End Function

    ' Método para insertar un hipervínculo en un control DataGridView
    Public Sub InsertarHipervinculo(celda As DataGridViewCell, nombreArchivo As String, Optional ext As String = "pdf")
        Dim rutaArchivo As String = BuscarArchivo(nombreArchivo, ext, SearchOption.AllDirectories)
        If String.IsNullOrEmpty(rutaArchivo) Then
            celda.Value = "No encontrado"
        Else
            celda.Value = "Abrir " & ext.ToUpper()
            celda.Tag = rutaArchivo ' Guarda la ruta en el Tag
        End If
    End Sub

    ' Método para abrir un archivo
    Public Sub AbrirArchivo(rutaArchivo As String)
        If File.Exists(rutaArchivo) Then
            Process.Start(New ProcessStartInfo With {
                .FileName = rutaArchivo,
                .UseShellExecute = True
            })
        Else
            MessageBox.Show("El archivo no existe: " & rutaArchivo, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
End Class

Public Class ContextMenuHandler
    Private contextMenu As ContextMenuStrip

    ' Constructor
    Public Sub New()
        contextMenu = New ContextMenuStrip()
        contextMenu.Items.Add("Abrir Archivo", Nothing, AddressOf AbrirArchivoDesdeMenu)
        contextMenu.Items.Add("Copiar Ruta", Nothing, AddressOf CopiarRutaDesdeMenu)
    End Sub

    ' Mostrar el menú contextual en la posición del cursor
    Public Sub MostrarMenu(rutaArchivo As String)
        If Not String.IsNullOrEmpty(rutaArchivo) Then
            contextMenu.Tag = rutaArchivo ' Almacena la ruta en el menú
            contextMenu.Show(Cursor.Position)
        End If
    End Sub

    ' Método para copiar la ruta del archivo al portapapeles
    Private Sub CopiarRutaDesdeMenu(sender As Object, e As EventArgs)
        If Not String.IsNullOrEmpty(contextMenu.Tag) Then
            Clipboard.SetText(contextMenu.Tag.ToString())
        End If
    End Sub

    ' Método para abrir el archivo desde el menú contextual
    Private Sub AbrirArchivoDesdeMenu(sender As Object, e As EventArgs)
        If Not String.IsNullOrEmpty(contextMenu.Tag) Then
            Dim pdfSearcher As New FileSearcher("")
            pdfSearcher.AbrirArchivo(contextMenu.Tag.ToString())
        End If
    End Sub
End Class
