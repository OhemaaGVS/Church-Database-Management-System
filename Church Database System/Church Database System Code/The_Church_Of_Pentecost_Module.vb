Imports System.Data.OleDb
Module The_Church_Of_Pentecost_Module
    Public Const DatabasePath As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='' ; Persist Security Info=False;" 'refers to the file path of the Access Database
    Public Connection As OleDbConnection
    Public Function DatabaseConnection() As Boolean 'can call this function whenever we need to open a connection to the database. It returns the result True or False depending on whether it opens the connection successfully.
        Try ' trying to see if there is a connection
            Connection = New OleDbConnection(DatabasePath) '(public) variable called Connection which will be used to refer to our connection to the database when we open it anywhere in the program.
            Connection.Open()
            Return True
        Catch ex As Exception
            MessageBox.Show("Unable to open the database :" & ex.Message) ' message box showing the data base can not be opend
            Return False
        End Try
    End Function
End Module


