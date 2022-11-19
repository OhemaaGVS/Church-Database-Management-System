
Imports System.Data.OleDb
Public Class Local_Ranking_Form
    Private Sub Local_Ranking_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' when this form loads, this sub procedure will display all the locals and rank them based on which local has the most members in the database
        Local_Rank_ListBox.Items.Clear()
        Local_Name_ListBox.Items.Clear()
        Local_Number_ListBox.Items.Clear() ' clearing the lit boxes
        If DatabaseConnection() Then ' checking for a connection 
            Dim Number As Integer = 0 ' this varialble holds the rank numbering for the ranking table 
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select * From LOCAL_TABLE Order By [number] Desc,LOCAL_NAME"
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim LocalNumber, LocalName As String ' these variables will hold the value that th record set reads
                    LocalNumber = DataReader("number")
                    LocalName = DataReader("LOCAL_NAME")
                    Number = Number + 1
                    Local_Rank_ListBox.Items.Add(Number)
                    Local_Name_ListBox.Items.Add(LocalName)
                    Local_Number_ListBox.Items.Add(LocalNumber) ' displaying the data in the list boxes
                End While
                DataReader.Close() ' close the the record set 
            End With
            Connection.Close() ' close the connection 
        End If
    End Sub
End Class