'Imports System.Data.SQLite
'Imports System.Windows.Forms

'Module SQLite
'    Private Sub LoadData(DataGridView As DataGridView)
'        Using conn As SQLiteConnection = DatabaseModule.GetConnection()
'            conn.Open()
'            Dim da As New SQLiteDataAdapter("SELECT * FROM MyTable", conn)
'            Dim ds As New DataSet()
'            da.Fill(ds)
'            Dim dt As System.Data.DataTable = ds.Tables(0)
'            DataGridView.DataSource = dt
'        End Using
'    End Sub
'    Private Sub SaveData(DataGridView As DataGridView)
'        Using conn As SQLiteConnection = DatabaseModule.GetConnection()
'            conn.Open()
'            Dim dt As DataTable = CType(DataGridView.DataSource, DataTable)
'            Dim da As New SQLiteDataAdapter("SELECT * FROM MyTable", conn)
'            Dim cb As New SQLiteCommandBuilder(da)
'            da.InsertCommand = cb.GetInsertCommand()
'            da.UpdateCommand = cb.GetUpdateCommand()
'            da.DeleteCommand = cb.GetDeleteCommand()
'            ' Ensure that the DataTable is part of a DataSet and has the correct name
'            Dim ds As New DataSet()
'            ds.Tables.Add() ' Copy the DataTable to avoid issues with naming
'            ds.Tables(0).TableName = "MyTable" ' Set the table name to match the database table name
'            da.Update(ds, "MyTable")
'        End Using
'    End Sub
'    Private Sub CreateTables()
'        Using conn As SQLiteConnection = DatabaseModule.GetConnection()
'            conn.Open()
'            Dim cmd As New SQLiteCommand With {
'                .Connection = conn,
'                .CommandText = "CREATE TABLE IF NOT EXISTS MyTable (ID INTEGER PRIMARY KEY AUTOINCREMENT, Name TEXT, Age INTEGER)"
'            }
'            cmd.ExecuteNonQuery()
'        End Using
'    End Sub
'End Module
