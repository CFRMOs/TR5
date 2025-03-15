Imports System.Data.SQLite

Public Class DatabaseHelper
    Private ReadOnly connectionString As String

    Public Sub New(databasePath As String)
        connectionString = $"Data Source={databasePath};Version=3;"
        ' Crea la base de datos si no existe
        CreateDatabaseIfNotExists(databasePath)
    End Sub

    Private Sub CreateDatabaseIfNotExists(databasePath As String)
        If Not System.IO.File.Exists(databasePath) Then
            SQLiteConnection.CreateFile(databasePath)
            Console.WriteLine("Base de datos creada en: " & databasePath)
            CrearTablaPolilineas()
        End If
    End Sub

    ' Crear una tabla si no existe
    Private Sub CrearTablaPolilineas()
        Using connection As New SQLiteConnection(connectionString)
            connection.Open()

            Dim query As String = "CREATE TABLE IF NOT EXISTS Polilineas (
                                    TabName TEXT PRIMARY KEY,
                                    Layer TEXT,
                                    MinX REAL,
                                    MinY REAL,
                                    MaxX REAL,
                                    MaxY REAL,
                                    StartStation REAL,
                                    EndStation REAL,
                                    Longitud REAL,
                                    Area REAL,
                                    Side TEXT,
                                    AlignmentHDI TEXT,
                                    Comentarios TEXT,
                                    Tramos TEXT,
                                    Plano TEXT,
                                    FileName TEXT,
                                    FilePath TEXT
                                  )"
            Dim command As New SQLiteCommand(query, connection)
            command.ExecuteNonQuery()

            Console.WriteLine("Tabla Polilineas creada o ya existe.")
        End Using
    End Sub
End Class
