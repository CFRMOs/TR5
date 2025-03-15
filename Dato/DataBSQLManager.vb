Imports System.Data.SQLite
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Excelinterop = Microsoft.Office.Interop
Imports Excel = Microsoft.Office.Interop.Excel

''' <summary>
''' Clase encargada de gestionar la base de datos SQLite para almacenar y gestionar datos dinámicamente.
''' </summary>
Public Class DataBSQLManager
    Public ReadOnly connectionString As String
    Public Property TableName As String

    Public dbFolderPath As String = ""

    ''' <summary>
    ''' Constructor que inicializa el gestor de la base de datos en la carpeta principal del proyecto.
    ''' </summary>
    Public Sub New()
        ' Obtener la ruta del directorio actual del proyecto (ruta de trabajo actual)
        Dim currentDir As String = Environment.CurrentDirectory 'Assembly.GetExecutingAssembly().Location 
        'Dim currentDir As String = Assembly.GetExecutingAssembly().Location

        ' Subir varios niveles hasta llegar a la carpeta raíz del proyecto
        Dim projectRoot As String = Path.GetFullPath(Path.Combine(currentDir, "..", "..", "..", ".."))

        ' Crear una subcarpeta para la base de datos si no existe
        dbFolderPath = Path.Combine(projectRoot, "Database")
        If Not Directory.Exists(dbFolderPath) Then
            Directory.CreateDirectory(dbFolderPath)
        End If

        ' Configurar la cadena de conexión sin crear automáticamente el archivo de base de datos
        Dim dbPath As String = Path.Combine(dbFolderPath, "CunetasData.sqlite")
        connectionString = $"Data Source={dbPath};Version=3;"
    End Sub

    ''' <summary>
    ''' Constructor que inicializa el gestor de la base de datos.
    ''' </summary>
    ''' <param name="dbPath">La ruta del archivo de la base de datos SQLite.</param>
    Public Sub New(dbPath As String)
        ' Obtener el directorio de la ruta de la base de datos
        Dim directory As String = Path.GetDirectoryName(dbPath)

        ' Verificar y crear el directorio si no existe
        If Not String.IsNullOrEmpty(directory) AndAlso Not System.IO.Directory.Exists(directory) Then
            System.IO.Directory.CreateDirectory(directory)
        End If

        ' Crear la base de datos si no existe
        If Not System.IO.File.Exists(dbPath) Then
            SQLiteConnection.CreateFile(dbPath)
        End If

        connectionString = $"Data Source={dbPath};Version=3;"
    End Sub
    ''' <summary>
    ''' Verifica si una tabla existe en la base de datos.
    ''' </summary>
    ''' <param name="tableName">El nombre de la tabla a verificar.</param>
    ''' <returns>Devuelve True si la tabla existe; de lo contrario, False.</returns>
    Private Function TableExists(tableName As String) As Boolean
        Using connection As New SQLiteConnection(connectionString)
            connection.Open()
            Using command As New SQLiteCommand($"SELECT name FROM sqlite_master WHERE type='table' AND name='{tableName}'", connection)
                Using reader As SQLiteDataReader = command.ExecuteReader()
                    Return reader.HasRows
                End Using
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Crea una tabla dinámica en la base de datos basada en las propiedades de una clase, si no existe.
    ''' </summary>
    ''' <typeparam name="T">El tipo de la clase cuyas propiedades se usarán para definir la tabla.</typeparam>
    ''' <param name="tableName">El nombre de la tabla a crear.</param>
    Public Sub CreateTable(Of T)(tableName As String)
        If TableExists(tableName) Then
            Console.WriteLine($"La tabla '{tableName}' ya existe. No se necesita crearla.")
            Return
        End If

        Dim props = GetType(T).GetProperties()
        Dim columns As New List(Of String)

        For Each prop As PropertyInfo In props
            Dim columnDefinition As String = $"{prop.Name} "
            If prop.PropertyType = GetType(String) Then
                columnDefinition &= "TEXT"
            ElseIf prop.PropertyType = GetType(Integer) Then
                columnDefinition &= "INTEGER"
            ElseIf prop.PropertyType = GetType(Double) OrElse prop.PropertyType = GetType(Single) Then
                columnDefinition &= "REAL"
            Else
                columnDefinition &= "TEXT"
            End If
            columns.Add(columnDefinition)
        Next

        columns.Insert(0, "Id INTEGER PRIMARY KEY AUTOINCREMENT")
        Dim query As String = $"CREATE TABLE {tableName} ({String.Join(", ", columns)})"

        Using connection As New SQLiteConnection(connectionString)
            connection.Open()
            Using command As New SQLiteCommand(query, connection)
                command.ExecuteNonQuery()
            End Using
        End Using

        Console.WriteLine($"Tabla '{tableName}' creada exitosamente.")
    End Sub

    ' Comprobación de si el Handle ya existe en la base de datos
    Public Function CheckForDuplicateHandle(tableName As String, handleValue As String) As Boolean
        If String.IsNullOrEmpty(handleValue) Then
            Console.WriteLine("Handle es NULL o vacío. No es necesario comprobar duplicados.")
            Return True ' Handle inválido, no proceder con la inserción
        End If

        Dim queryCheck As String = $"SELECT COUNT(*) FROM {tableName} WHERE Handle = @Handle"
        Using connection As New SQLiteConnection(connectionString)
            connection.Open()
            Using command As New SQLiteCommand(queryCheck, connection)
                command.Parameters.AddWithValue("@Handle", handleValue)
                Dim count As Integer = Convert.ToInt32(command.ExecuteScalar())
                If count > 0 Then
                    Console.WriteLine($"Ya existe un registro con Handle: {handleValue}. La inserción se omite.")
                    Return True ' Duplicado encontrado
                End If
            End Using
        End Using
        Return False ' No se encontraron duplicados
    End Function

    ' Inserción de un objeto en la tabla de la base de datos
    Public Sub InsertDataItem(Of T)(dataItem As T, tableName As String)
        Dim props = dataItem.GetType().GetProperties()

        Dim query As String = $"INSERT INTO {tableName} ("
        Dim columnNames As New List(Of String)
        Dim parameterNames As New List(Of String)

        For Each prop In props
            columnNames.Add(prop.Name)
            parameterNames.Add("@" & prop.Name)
        Next

        query &= String.Join(", ", columnNames) & ") VALUES (" & String.Join(", ", parameterNames) & ")"

        Using connection As New SQLiteConnection(connectionString)
            connection.Open()
            Using command As New SQLiteCommand(query, connection)
                For Each prop In props
                    Dim value = prop.GetValue(dataItem)
                    If value Is Nothing Then
                        command.Parameters.AddWithValue("@" & prop.Name, DBNull.Value)
                    Else
                        command.Parameters.AddWithValue("@" & prop.Name, value)
                    End If
                Next
                command.ExecuteNonQuery()
                Console.WriteLine("Datos insertados exitosamente.")
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Comprueba si ya existe un registro en la base de datos con las propiedades especificadas de un objeto CunetasHandleDataItem.
    ''' </summary>
    ''' <param name="tableName">El nombre de la tabla donde se realizará la comprobación.</param>
    ''' <param name="dataItem">El objeto CunetasHandleDataItem con las propiedades a comparar.</param>
    ''' <returns>Devuelve True si se encuentra un registro duplicado; de lo contrario, False.</returns>
    Public Function CheckForDuplicateEntry(tableName As String, dataItem As CunetasHandleDataItem) As Boolean
        Dim tolerancia As Double = 2.0
        Dim precision As Integer = 2

        ' Construir la consulta SQL con las condiciones
        Dim queryCheck As String = $"SELECT COUNT(*) FROM {tableName} WHERE " &
                               "Accesos = @Accesos AND " &
                               "ABS(ROUND(StartStation, @Precision) - ROUND(@StartStation, @Precision)) <= @Tolerancia AND " &
                               "ABS(ROUND(EndStation, @Precision) - ROUND(@EndStation, @Precision)) <= @Tolerancia AND " &
                               "ABS(ROUND(Longitud, @Precision) - ROUND(@Longitud, @Precision)) <= @Tolerancia AND " &
                               "Side = @Side"

        ' Agregar la condición de IDNum solo si es distinto de 0
        If dataItem.IDNum <> 0 Then
            queryCheck &= " AND IDNum = @IDNum"
        End If

        Using connection As New SQLiteConnection(connectionString)
            connection.Open()
            Using command As New SQLiteCommand(queryCheck, connection)
                ' Asignar los parámetros de la consulta
                command.Parameters.AddWithValue("@Accesos", dataItem.Accesos)
                command.Parameters.AddWithValue("@StartStation", dataItem.StartStation)
                command.Parameters.AddWithValue("@EndStation", dataItem.EndStation)
                command.Parameters.AddWithValue("@Longitud", dataItem.Longitud)
                command.Parameters.AddWithValue("@Side", dataItem.Side)
                command.Parameters.AddWithValue("@Tolerancia", tolerancia)
                command.Parameters.AddWithValue("@Precision", precision)

                If dataItem.IDNum <> 0 Then
                    command.Parameters.AddWithValue("@IDNum", dataItem.IDNum)
                End If

                Dim count As Integer = Convert.ToInt32(command.ExecuteScalar())
                If count > 0 Then
                    Console.WriteLine("Se encontró un registro duplicado basado en las propiedades especificadas.")
                    Return True ' Duplicado encontrado
                End If
            End Using
        End Using

        Return False ' No se encontraron duplicados
    End Function

    ' Método principal que usa las dos partes: comprobación e inserción
    ''' <summary>
    ''' Inserta un objeto en la base de datos, comprobando primero si el Handle es nulo o duplicado
    ''' y si las propiedades adicionales representan un registro existente.
    ''' </summary>
    ''' <param name="dataItem">El objeto CunetasHandleDataItem con los datos que se deben insertar.</param>
    ''' <param name="tableName">El nombre de la tabla en la que se realizará la inserción.</param>
    Public Sub InsertData(dataItem As CunetasHandleDataItem, tableName As String)
        ' Comprobar si el Handle es NULL o si ya existe en la base de datos
        If String.IsNullOrEmpty(dataItem.Handle) OrElse CheckForDuplicateHandle(tableName, dataItem.Handle) Then
            Console.WriteLine("El Handle es nulo o ya existe. La inserción se omite.")
            Return ' No proceder con la inserción si el Handle es nulo o duplicado
        End If

        ' Comprobar si ya existe un registro con las propiedades adicionales
        If CheckForDuplicateEntry(tableName, dataItem) Then
            Console.WriteLine("Se encontró un registro duplicado basado en las propiedades especificadas. La inserción se omite.")
            Return ' No proceder con la inserción si se encontró un duplicado
        End If

        ' Proceder con la inserción si no se encontraron duplicados
        InsertDataItem(dataItem, tableName)
        Console.WriteLine("Datos insertados exitosamente.")
    End Sub

    ''' <summary>
    ''' Crea una tabla en la base de datos si no existe, incluyendo un campo de ID personalizado como clave primaria.
    ''' </summary>
    ''' <typeparam name="T">El tipo de la clase que define la estructura de la tabla.</typeparam>
    ''' <param name="tableName">El nombre de la tabla.</param>
    ''' <param name="primaryKeyName">El nombre del campo de ID que actuará como clave primaria autoincremental.</param>
    Public Sub CreateTableIfNotExists(Of T)(tableName As String, primaryKeyName As String)
        ' Validar que el nombre del ID no esté vacío
        If String.IsNullOrEmpty(primaryKeyName) Then
            Throw New ArgumentException("El nombre de la clave primaria no puede estar vacío.", NameOf(primaryKeyName))
        End If

        ' Obtener las propiedades de la clase T
        Dim props = GetType(T).GetProperties()
        Dim columns As New List(Of String)

        ' Añadir la columna de ID personalizado como clave primaria autoincremental
        columns.Add($"{primaryKeyName} INTEGER PRIMARY KEY AUTOINCREMENT")

        ' Generar las definiciones de columnas basadas en los tipos de propiedades
        For Each prop As PropertyInfo In props
            Dim columnDefinition As String = $"{prop.Name} "

            ' Determinar el tipo de dato SQLite según el tipo de propiedad
            If prop.PropertyType = GetType(String) Then
                columnDefinition &= "TEXT"
            ElseIf prop.PropertyType = GetType(Integer) Then
                columnDefinition &= "INTEGER"
            ElseIf prop.PropertyType = GetType(Double) OrElse prop.PropertyType = GetType(Single) Then
                columnDefinition &= "REAL"
            Else
                columnDefinition &= "TEXT" ' Tipo de datos por defecto
            End If

            columns.Add(columnDefinition)
        Next

        ' Crear la consulta de creación de tabla
        Dim query As String = $"CREATE TABLE IF NOT EXISTS {tableName} ({String.Join(", ", columns)})"

        ' Ejecutar la consulta para crear la tabla
        Using connection As New SQLiteConnection(connectionString)
            connection.Open()
            Using command As New SQLiteCommand(query, connection)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub


    ''' <summary>
    ''' Obtiene un objeto de la tabla por el Handle.
    ''' </summary>
    ''' <typeparam name="T">El tipo del objeto a obtener.</typeparam>
    ''' <param name="tableName">El nombre de la tabla de la que se obtiene el objeto.</param>
    ''' <param name="handle">El Handle del objeto a buscar.</param>
    ''' <returns>Devuelve una instancia del objeto si se encuentra; de lo contrario, Nothing.</returns>
    Public Function GetDataItemByHandle(Of T As {Class, New})(tableName As String, handle As String) As T
        Try
            Dim query As String = $"SELECT * FROM {tableName} WHERE Handle = @Handle"
            Using connection As New SQLiteConnection(connectionString)
                connection.Open()
                Using command As New SQLiteCommand(query, connection)
                    command.Parameters.AddWithValue("@Handle", handle)
                    Using reader As SQLiteDataReader = command.ExecuteReader()
                        If reader.Read() Then
                            Dim dataItem As T = Activator.CreateInstance(Of T)()
                            Dim props = dataItem.GetType().GetProperties()
                            For Each prop In props
                                ' Skip properties that don't have a Set method (read-only properties)
                                If Not prop.CanWrite Then
                                    Continue For
                                End If
                                If Not IsDBNull(reader(prop.Name)) Then
                                    Dim value As Object = reader(prop.Name)
                                    ' Handle special cases for complex types
                                    If prop.PropertyType Is GetType(List(Of Integer)) Then
                                        ' Example: Convert a comma-separated string to a List(Of Integer)
                                        Dim intList As New List(Of Integer)
                                        Dim stringValue As String = value.ToString()
                                        If Not String.IsNullOrEmpty(stringValue) Then
                                            ' Safely parse each part of the string
                                            For Each part As String In stringValue.Split(","c)
                                                Dim number As Integer
                                                If Integer.TryParse(part.Trim(), number) Then
                                                    intList.Add(number)
                                                Else
                                                    ' Log or handle invalid numbers
                                                    Console.WriteLine($"Invalid integer value: '{part}'")
                                                End If
                                            Next
                                        End If
                                        prop.SetValue(dataItem, intList)
                                    Else
                                        ' Use Convert.ChangeType for primitive types
                                        prop.SetValue(dataItem, Convert.ChangeType(value, prop.PropertyType))
                                    End If
                                End If
                            Next
                            Return dataItem
                        End If
                    End Using
                End Using
            End Using
        Catch ex As SQLiteException
            Console.WriteLine($"SQLiteException: {ex.Message}")
        Catch ex As InvalidCastException
            Console.WriteLine($"InvalidCastException: {ex.Message}")
        Catch ex As ArgumentException
            Console.WriteLine($"ArgumentException: {ex.Message}")
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Obtiene el ID de la clave primaria de la fila de la base de datos por el Handle.
    ''' </summary>
    ''' <param name="tableName">El nombre de la tabla.</param>
    ''' <param name="dataItem">El objeto CunetasHandleDataItem con las propiedades a comparar.</param>
    ''' <param name="primaryKeyName">El nombre del campo de la clave primaria.</param>
    ''' <returns>Devuelve el ID de la clave primaria si se encuentra, o -1 si no se encuentra.</returns>
    Public Function GetRowByHandle(tableName As String, dataItem As CunetasHandleDataItem, primaryKeyName As String) As Integer
        If String.IsNullOrEmpty(dataItem.Handle) Then Return -1

        Dim query As String = $"SELECT {primaryKeyName} FROM {tableName} WHERE Handle = @Handle LIMIT 1"
        Using connection As New SQLiteConnection(connectionString)
            connection.Open()
            Using command As New SQLiteCommand(query, connection)
                command.Parameters.AddWithValue("@Handle", dataItem.Handle)
                Dim result As Object = command.ExecuteScalar()
                If result IsNot Nothing Then
                    Return Convert.ToInt32(result)
                End If
            End Using
        End Using

        Return -1 ' No se encontró el registro
    End Function

    ''' <summary>
    ''' Obtiene el ID de la clave primaria de la fila de la base de datos por las propiedades del objeto con tolerancia.
    ''' </summary>
    ''' <param name="tableName">El nombre de la tabla.</param>
    ''' <param name="dataItem">El objeto CunetasHandleDataItem con las propiedades a comparar.</param>
    ''' <param name="primaryKeyName">El nombre del campo de la clave primaria.</param>
    ''' <returns>Devuelve el ID de la clave primaria si se encuentra, o -1 si no se encuentra.</returns>
    Public Function GetRowByProperties(tableName As String, dataItem As CunetasHandleDataItem, primaryKeyName As String) As Integer
        Dim tolerancia As Double = 2.0
        Dim precision As Integer = 2
        Dim query As String = $"SELECT {primaryKeyName} FROM {tableName} WHERE " &
                          "Accesos = @Accesos AND " &
                          "ABS(ROUND(StartStation, @Precision) - ROUND(@StartStation, @Precision)) <= @Tolerancia AND " &
                          "ABS(ROUND(EndStation, @Precision) - ROUND(@EndStation, @Precision)) <= @Tolerancia AND " &
                          "ABS(ROUND(Longitud, @Precision) - ROUND(@Longitud, @Precision)) <= @Tolerancia AND " &
                          "Side = @Side LIMIT 1"

        Using connection As New SQLiteConnection(connectionString)
            connection.Open()
            Using command As New SQLiteCommand(query, connection)
                command.Parameters.AddWithValue("@Accesos", dataItem.Accesos)
                command.Parameters.AddWithValue("@StartStation", dataItem.StartStation)
                command.Parameters.AddWithValue("@EndStation", dataItem.EndStation)
                command.Parameters.AddWithValue("@Longitud", dataItem.Longitud)
                command.Parameters.AddWithValue("@Side", dataItem.Side)
                command.Parameters.AddWithValue("@Tolerancia", tolerancia)
                command.Parameters.AddWithValue("@Precision", precision)

                Dim result As Object = command.ExecuteScalar()
                If result IsNot Nothing Then
                    Return Convert.ToInt32(result)
                End If
            End Using
        End Using

        Return -1 ' No se encontró el registro
    End Function

    ''' <summary>
    ''' Actualiza una fila en la base de datos utilizando el ID de la clave primaria.
    ''' </summary>
    ''' <param name="tableName">El nombre de la tabla.</param>
    ''' <param name="dataItem">El objeto CunetasHandleDataItem con los datos actualizados.</param>
    ''' <param name="primaryKeyName">El nombre del campo de la clave primaria.</param>
    Public Sub UpdateDataID(tableName As String, dataItem As CunetasHandleDataItem, primaryKeyName As String, id As Integer)
        If id = -1 Then
            Console.WriteLine("No se proporcionó un ID válido para actualizar.")
            Return
        End If

        ' Construir la consulta UPDATE usando reflexión
        Dim query As String = $"UPDATE {tableName} SET "
        Dim props = dataItem.GetType().GetProperties()
        Dim setClauses As New List(Of String)

        ' Construir las cláusulas SET excluyendo la clave primaria
        For Each prop In props
            If prop.Name <> primaryKeyName Then
                setClauses.Add($"{prop.Name} = @{prop.Name}")
            End If
        Next

        query &= String.Join(", ", setClauses) & $" WHERE {primaryKeyName} = @{primaryKeyName}"

        ' Ejecutar la consulta SQL
        Using connection As New SQLiteConnection(connectionString)
            connection.Open()
            Using command As New SQLiteCommand(query, connection)
                ' Agregar los parámetros para todas las propiedades del objeto
                For Each prop In props
                    Dim value = prop.GetValue(dataItem)
                    If value Is Nothing Then
                        command.Parameters.AddWithValue("@" & prop.Name, DBNull.Value)
                    Else
                        command.Parameters.AddWithValue("@" & prop.Name, value)
                    End If
                Next

                ' Agregar el parámetro para la clave primaria
                command.Parameters.AddWithValue($"@{primaryKeyName}", id)

                ' Ejecutar la consulta
                command.ExecuteNonQuery()
            End Using
        End Using

        Console.WriteLine("Registro actualizado exitosamente.")
    End Sub

    ''' <summary>
    ''' Obtiene un valor desde la base de datos basado en el ID y el nombre de la columna.
    ''' </summary>
    ''' <param name="tableName">El nombre de la tabla.</param>
    ''' <param name="id">El ID del registro a buscar.</param>
    ''' <param name="primaryKeyName">El nombre de la columna que representa la clave primaria.</param>
    ''' <param name="columnName">El nombre de la columna cuyo valor se desea obtener.</param>
    ''' <returns>El valor de la columna como un objeto, o Nothing si no se encuentra.</returns>
    Public Function GetDataById(tableName As String, id As Integer, primaryKeyName As String, columnName As String) As Object
        Dim value As Object = Nothing

        Using connection As New SQLiteConnection(connectionString)
            connection.Open()
            Dim query As String = $"SELECT {columnName} FROM {tableName} WHERE {primaryKeyName} = @ID"
            Using command As New SQLiteCommand(query, connection)
                command.Parameters.AddWithValue("@ID", id)
                Using reader As SQLiteDataReader = command.ExecuteReader()
                    If reader.Read() Then
                        value = reader(columnName)
                    End If
                End Using
            End Using
        End Using

        Return value
    End Function



    ''' <summary>
    ''' Actualiza un objeto en la tabla de la base de datos.
    ''' </summary>
    ''' <typeparam name="T">El tipo del objeto a actualizar.</typeparam>
    ''' <param name="dataItem">El objeto con los datos actualizados.</param>
    ''' <param name="tableName">El nombre de la tabla en la que se actualizará el objeto.</param>
    Public Sub UpdateData(Of T)(dataItem As T, tableName As String)
        Dim query As String = $"UPDATE {tableName} SET "
        Dim props = dataItem.GetType().GetProperties()
        Dim setClauses As New List(Of String)

        For Each prop In props
            If prop.Name <> "Handle" Then
                setClauses.Add($"{prop.Name} = @{prop.Name}")
            End If
        Next

        query &= String.Join(", ", setClauses) & " WHERE Handle = @Handle"

        Using connection As New SQLiteConnection(connectionString)
            connection.Open()
            Using command As New SQLiteCommand(query, connection)
                For Each prop In props
                    Dim value = prop.GetValue(dataItem)
                    If value Is Nothing Then
                        command.Parameters.AddWithValue("@" & prop.Name, DBNull.Value)
                    Else
                        command.Parameters.AddWithValue("@" & prop.Name, value)
                    End If
                Next
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub




    ''' <summary>
    ''' Método que busca objetos en la base de datos por longitud, con una tolerancia especificada.
    ''' </summary>
    ''' <param name="tableName">El nombre de la tabla donde se realizará la búsqueda.</param>
    ''' <param name="longitud">La longitud a buscar.</param>
    ''' <param name="tolerancia">La tolerancia de búsqueda para la longitud (por defecto es 2.0).</param>
    ''' <returns>Devuelve una lista de objetos CunetasHandleDataItem que cumplen con el criterio de búsqueda.</returns>
    Public Function SearchByLengthInDatabase(ByVal tableName As String,
                                             ByVal longitud As Double,
                                             Optional ByVal tolerancia As Double = 2.0) As List(Of CunetasHandleDataItem)

        Dim resultados As New List(Of CunetasHandleDataItem)

        ' Construir la consulta SQL con una tolerancia para la longitud
        Dim query As String = $"SELECT * FROM {tableName} WHERE ABS(Longitud - @Longitud) <= @Tolerancia"

        Using connection As New SQLiteConnection(connectionString)
            connection.Open()
            Using command As New SQLiteCommand(query, connection)
                ' Asignar parámetros a la consulta
                command.Parameters.AddWithValue("@Longitud", longitud)
                command.Parameters.AddWithValue("@Tolerancia", tolerancia)

                Using reader As SQLiteDataReader = command.ExecuteReader()
                    ' Leer los resultados y convertirlos en objetos CunetasHandleDataItem
                    While reader.Read()
                        Dim dataItem As New CunetasHandleDataItem()
                        For Each prop In dataItem.GetType().GetProperties()
                            If Not IsDBNull(reader(prop.Name)) Then
                                Dim value As Object = reader(prop.Name)
                                prop.SetValue(dataItem, Convert.ChangeType(value, prop.PropertyType))
                            End If
                        Next
                        resultados.Add(dataItem)
                    End While
                End Using
            End Using
        End Using

        Return resultados
    End Function
End Class

Public Class UTLBaseDAtos
    Public EventsHandler As New List(Of EventHandlerClass)
    Public Sub New(TabPanol As TabControl, TabCtrlViews As TabControl, CAcadHelp As ACAdHelpers, ComboBox1 As ComboBox, ComMediciones As ComboBox, handleProcessor As HandleDataProcessor)
        'Dim CunetasExistentesDGView As DataGridView = GetControlsDGView.GetDGViewByTabName(TabCtrlViews, "")
        Dim Panel As New TabsMenu(TabPanol) With {
            .TabName = "Base de Datos",
            .ButtonActions = New Dictionary(Of String, Action) From {
                                                                        {"Cambiar Layer", Sub() CtrolsButtonsUT.CambiarLayer(ComboBox1, ComMediciones, TabCtrlViews, CAcadHelp)},
                                                                        {"Selected By Entity Soporte", Sub() SelectedByentitySoporte(TabCtrlViews)}
                                                                    }
                                                 }


    End Sub
    Public Sub SelectedByentitySoporte(Optional TabCtrolView As TabControl = Nothing)

        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

        Dim DGView As DataGridView = GetControlsDGView.GetDGView(TabCtrolView)

        Dim iHeaders As List(Of String) = HandleDataProcessor.Headers.Select(Function(item) item.Item1).ToList()

        Dim cListType As New List(Of List(Of Autodesk.AutoCAD.DatabaseServices.ObjectId))

        ' Construct the dictionary only once
        Dim ListType As List(Of List(Of Autodesk.AutoCAD.DatabaseServices.ObjectId)) = AutoCADHelper2.ConstructDiccionario(cListType)

        'If EncabezadosSoporte Is Nothing Then
        Dim ExRNGSe As New ExcelRangeSelector()
            ' Select the range and search for headers
            Dim RNGS As Excel.Range = ExRNGSe.SelectRange()
            Using optimizer As New ExcelOptimizer(ExRNGSe.ExcelApp.GetWorkB().Application)
                Try
                    optimizer.TurnEverythingOff()

                    ' Loop through the Excel range
                    For Each Rng As Excel.Range In RNGS
                        Try
                            ' Create a new CunetasHandleDataItem and populate properties
                            Dim Cuneta As New CunetasHandleDataItem

                            SetProFromSoporteCuneta(Cuneta, Rng)

                            ' Handle observations
                            Dim RNGOBSERVACION As Excel.Range = RNGOffset(Rng, "OBSERVACIONES TYPSA")

                            ' Sync the data and update the DataGridView
                            SyncSoporteTabla(Cuneta, Cuneta.IDNum, doc, ExRNGSe, RNGOBSERVACION, ListType)

                            ' Only add to DGView if there's no observation value
                            If RNGOBSERVACION.Value = "" Then
                                Cuneta.AddToDGView(DGView)
                            End If
                        Catch ex As Exception
                            ' Log or handle errors specific to this row
                            Debug.WriteLine($"Error processing row {Rng.Address}: {ex.Message}")
                        End Try
                    Next
                    ' Explicitly release the Excel COM objects
                    Marshal.ReleaseComObject(RNGS)
                Catch ex As Exception
                    ' Handle overall errors here
                    Debug.WriteLine($"Error in SelectedByentitySoporte: {ex.Message}")
                Finally
                    optimizer.TurnEverythingOn()
                    ExRNGSe = Nothing
                End Try
            End Using
        'End If

    End Sub
    Public Sub SetProFromSoporteCuneta(ByRef Cuneta As CunetasHandleDataItem, Rng As Excel.Range)

        Dim STATIONS As New List(Of Double) From {CDbl(RNGOffset(Rng, "KM INCIAL").Value), CDbl(RNGOffset(Rng, "KM FINAL").Value)}

        Cuneta.StartStation = STATIONS.Min

        Cuneta.EndStation = STATIONS.Max

        Cuneta.Side = If(RNGOffset(Rng, "LADO").Value.ToString() = "DER", "Right", "Left")

        Cuneta.Longitud = Math.Abs(CDbl(RNGOffset(Rng, "LONG.(m)").Value))

        Cuneta.IDNum = CDbl(Replace(RNGOffset(Rng, "ID").Value, "CU-", ""))
    End Sub
    Public Function RNGOffset(RNG As Excel.Range, Header As String) As Excel.Range
        Return RNG.Offset(0, RNG.Worksheet.Cells.Find(What:=Header, MatchCase:=True, LookAt:=Excel.XlLookAt.xlWhole).Column - RNG.Column)
    End Function
    Public Sub SyncSoporteTabla(Cuneta As CunetasHandleDataItem, IDNum As Double, doc As Document, ExRNGSe As ExcelRangeSelector, ByRef RNGnotM As Excel.Range, Optional ByRef ListType As List(Of List(Of Autodesk.AutoCAD.DatabaseServices.ObjectId)) = Nothing)
        ' Diccionarios para manejar los handles y los resultados
        Dim DicHandles As New Dictionary(Of String, CunetasHandleDataItem)
        Dim DicHandlesResultado As New Dictionary(Of String, CunetasHandleDataItem)

        ' Inicializar la variable para determinar si se ha encontrado un registro
        Dim found As Boolean = False
        Dim id As Integer = -1

        ' Intentar buscar primero por propiedades en la base de datos y obtener el ID
        'id = Cuneta.DbManager.GetRowByProperties("CunetasDatos", Cuneta, "ID_Cunetas")
        'If id <> -1 Then
        '    found = True
        'Else
        '    id = Cuneta.DbManager.GetRowByHandle("CunetasDatos", Cuneta, "ID_Cunetas")
        '    If id <> -1 Then
        '        found = True
        '    End If
        'End If

        ' Si se encuentra un resultado
        If found Then
            ' Obtener el layer desde la base de datos utilizando el ID y primaryKeyName
            'Dim layerFromDB As String = Cuneta.DbManager.GetDataById("CunetasDatos", id, "ID_Cunetas", "Layer").ToString()



            ' Configurar las propiedades desde el archivo DWG
            Cuneta.SetPropertiesFromDWG(Cuneta.Handle, Cuneta.Alignment.Handle.ToString())

            ' Asignar el número de ID
            Cuneta.IDNum = IDNum

            ' Asignar el layer desde la base de datos
            'Cuneta.Layer = layerFromDB

            ' Actualizar Cuneta en la base de datos
            Cuneta.SyncWithDatabase()

            ' Limpiar el valor en la celda de Excel
            RNGnotM.Value = ""
        Else
            ' Si no se encuentra, realizar la búsqueda por longitud
            DicHandles.Add("Not", Cuneta)
            'AutoCADHelper2.LookByLength(DicHandles, DicHandlesResultado, cListType:=ListType)

            ' Si se encuentran resultados por longitud
            If DicHandlesResultado.Count <> 0 Then
                Cuneta = DicHandlesResultado.Values(0)
                ' Obtener el layer de la tabla antes de actualizar las propiedades del DWG
                'Dim layerFromDB As String = Cuneta.DbManager.GetDataById("CunetasDatos", id, "ID_Cunetas", "Layer").ToString()
                'Cuneta.Layer = layerFromDB

                ' Configurar las propiedades desde el archivo DWG
                Cuneta.SetPropertiesFromDWG(Cuneta.Handle, Cuneta.Alignment.Handle.ToString())
                Cuneta.IDNum = IDNum

                ' Actualizar Cuneta en la base de datos
                Cuneta.SyncWithDatabase()

                RNGnotM.Value = ""
            Else
                ' Si aún no se encuentra nada, muestra el mensaje
                RNGnotM.Value = "No se encontró en la base de datos"
            End If
        End If
    End Sub

    Public Sub UpDateByExcelCuneta(CUNETA As CunetasHandleDataItem, Optional RECALCULATE As Boolean = False)

        Dim ExRNGSe As New ExcelRangeSelector()

        Dim TB As Excel.ListObject = ExRNGSe.GetTableOnWorkBkByName("CunetasGeneral")

        If RECALCULATE Then
            Dim AligmentsCertify As New AlignmentHelper()

            AligmentsCertify.CheckAlimentByProximida(CLHandle.GetEntityByStrHandle(CUNETA.Handle))

            CUNETA.SetPropertiesFromDWG(CUNETA.Handle, AligmentsCertify.Alignment?.Handle.ToString())
        End If

        Dim optimizer As New ExcelOptimizer(ExRNGSe.ExcelApp.GetWorkB().Application)

        optimizer.TurnEverythingOff()

        'proceso para añadir datos a excel 
        Dim Exp As New List(Of String) From {"Comentarios", "Ubicacion", "Medicion de Pago"}

        'CUNETA.Accesos = 9
        Dim RNG As Microsoft.Office.Interop.Excel.Range

        With CUNETA
            RNG = GetRNGByLeght(TB, .Accesos, .StartStation, .EndStation, .Longitud, .Side)
            If RNG Is Nothing Then RNG = GetRNGByHandle(CUNETA.Handle, TB)
        End With

        Dim Formatos As List(Of String) = HandleDataProcessor.Headers.Select(Function(item) item.Item3).ToList()

        CUNETA.AddToExcelTable(TB, Exp, RNG, Formatos)

        optimizer.TurnEverythingOn()
        optimizer.Dispose()
    End Sub
    Function GetRNGByIDNUM(ID As Integer, Acceso As Integer, Table As Microsoft.Office.Interop.Excel.ListObject) As Microsoft.Office.Interop.Excel.Range

        'buscar a partir de un Handle en la cloumna Handle
        Dim ExRNGSe As New ExcelRangeSelector()

        Dim indexHeader As Integer() = UpdateExcelTable.ConArray({"IDNum", "Accesos"}, Table.HeaderRowRange)

        Dim IDColumnIndex As Integer = indexHeader(0)

        Dim AccesoColumnIndex As Integer = indexHeader(1)

        Dim RNG As Microsoft.Office.Interop.Excel.Range = Nothing

        For Each row As Microsoft.Office.Interop.Excel.ListRow In Table.ListRows
            If row.Range.Cells(1, IDColumnIndex).Value = ID AndAlso row.Range.Cells(1, AccesoColumnIndex).Value = Acceso Then
                RNG = row.Range ' Return the matching row if found
                Exit For
            End If
        Next
        Return RNG
    End Function
    Function GetRNGByHandle(handle As String, Table As Microsoft.Office.Interop.Excel.ListObject) As Microsoft.Office.Interop.Excel.Range

        'buscar a partir de un Handle en la cloumna Handle
        Dim ExRNGSe As New ExcelRangeSelector()

        Dim indexHeader As Integer() = UpdateExcelTable.ConArray({"Handle", "Comentarios"}, Table.HeaderRowRange)

        Dim handleColumnIndex As Integer = indexHeader(0)

        Dim ComentariosColumnIndex As Integer = indexHeader(1)

        Dim RNG As Microsoft.Office.Interop.Excel.Range = Nothing

        For Each row As Microsoft.Office.Interop.Excel.ListRow In Table.ListRows
            If row.Range.Cells(1, handleColumnIndex).Value.ToString() = handle Then
                RNG = row.Range ' Return the matching row if found
                Exit For
            End If
        Next
        Return RNG
    End Function
    Function GetRNGByLeght(Table As Microsoft.Office.Interop.Excel.ListObject,
                           Acceso As Integer,
                           Startstation As Double,
                           Endstation As Double,
                           Longitud As Double,
                           Side As String, Optional ByRef IDNum As Integer = 0) As Microsoft.Office.Interop.Excel.Range

        'buscar a partir de un Handle en la cloumna Handle
        Dim ExRNGSe As New ExcelRangeSelector()
        Dim Tolerancia As Double = 2
        Dim Presicion As Integer = 2


        Dim indexHeader As Integer() = UpdateExcelTable.ConArray({"Accesos", "Start Station", "End Station", "Longitud", "Side", "IDNum"}, Table.HeaderRowRange)

        Dim RNG As Microsoft.Office.Interop.Excel.Range = Nothing

        For Each row As Microsoft.Office.Interop.Excel.ListRow In Table.ListRows
            If row.Range.Cells(1, indexHeader(0)).Value = Acceso AndAlso
              Math.Abs(Math.Round(CDbl(row.Range.Cells(1, indexHeader(1)).Value), Presicion) - Startstation) <= Tolerancia AndAlso
              Math.Abs(Math.Round(CDbl(row.Range.Cells(1, indexHeader(2)).Value), Presicion) - Endstation) <= Tolerancia AndAlso
               Math.Abs(Math.Round(CDbl(row.Range.Cells(1, indexHeader(3)).Value), Presicion) - Longitud) <= Tolerancia AndAlso
               row.Range.Cells(1, indexHeader(4)).Value.ToString() = Side Then

                Dim EvaNum As Boolean = True

                If IDNum <> 0 Then EvaNum = IDNum = row.Range.Cells(1, indexHeader(5)).Value

                If EvaNum Then
                    RNG = row.Range ' Return the matching row if found
                    Exit For
                End If
            End If
        Next
        Return RNG
    End Function
End Class


