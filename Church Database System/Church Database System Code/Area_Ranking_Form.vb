
Imports System.Data.OleDb
Public Class Area_Ranking_Form
    ' when this form loads, this sub procedure will display all the areas and rank them based on which area has the most districts in the database
    Private Sub Area_Ranking_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Rank_ListBox.Items.Clear()
        Area_ListBox.Items.Clear()
        Number_ListBox.Items.Clear() ' clearing the list boxes
        If DatabaseConnection() Then ' checking for a connection 
            Dim Number As Integer = 0 ' this varialble holds the rank numbering for the ranking table 
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select * From AREA_TABLE Order By [number] Desc,AREA_NAME"

                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim AreaNumber, AreaName As String ' these variables will hold the value that th record set reads
                    AreaNumber = DataReader("number")
                    AreaName = DataReader("AREA_NAME")
                    Number = Number + 1
                    Rank_ListBox.Items.Add(Number)
                    Area_ListBox.Items.Add(AreaName)
                    Number_ListBox.Items.Add(AreaNumber) ' displaying the data in the list boxes
                End While
                DataReader.Close() ' close the the record set
            End With
            Connection.Close() ' close the connection 

        End If
    End Sub
End Class