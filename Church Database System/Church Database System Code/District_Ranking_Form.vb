
Imports System.Data.OleDb
Public Class District_Ranking_Form
    ' when this form loads, this sub procedure will display all the Districts and rank them based on which district has the most locals in the database
    Private Sub District_Ranking_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Rank_ListBox.Items.Clear()
        District_ListBox.Items.Clear()
        Number_ListBox.Items.Clear() ' clearing the list boxes
        If DatabaseConnection() Then ' checking for a connection 
            Dim Number As Integer = 0 ' this varialble holds the rank numbering for the ranking table 
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select * From DISTRICT_TABLE Order By [number] Desc,DISTRICT_NAME"

                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim DistrictNumber, DistrictName As String ' these variables will hold the value that th record set reads
                    DistrictNumber = DataReader("number")
                    DistrictName = DataReader("DISTRICT_NAME")
                    Number = Number + 1
                    Rank_ListBox.Items.Add(Number)
                    District_ListBox.Items.Add(DistrictName)
                    Number_ListBox.Items.Add(DistrictNumber) ' displaying the data in the list boxes
                End While
                DataReader.Close() ' close the the record set
            End With
            Connection.Close() ' close the connection

        End If
    End Sub
End Class