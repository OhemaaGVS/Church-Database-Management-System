Imports System.Data.OleDb
Public Class Area_Management_Form
    Private ValueOfAreaID As Integer = -1 ' this variable will store the value of the area id that is generated for the new area that is being created 
    Private Area_Head_ID As Integer 'this variable will store the value of the initial selected area head for the area
    Private New_Area_Head_ID As Integer ' this will sotre the value of the new area head id when an area is being updated 
    Private Selected_Area_ID As Integer ' this will store the value of the area id that is selected form the list box by the user
    Private New_Area_ID As Integer ' this will store the new area id that the districts under the old area id will be given when an area is going to be deleted 
    Private RemoveNumberInAreas As Integer ' this variable will store the number of districts that was within the area that is about to be deleted 
    Private TemporalAreaNumber As Integer ' this variable will store the current number of districts within the area that will temporarily hold the data of the deleted area 
    Private CurrentYearsActive As Integer ' this variable will store the amount of years an area head has been when the area was first created in the database
    Private UnAssignedHeadID As Integer 'this variable will store the area head id of the area head who will become un asigned
    Private NewYearsActive As Integer ' this will store the new amount of years that the area head has been  based at an area when the area is being updated 
    Private FirstAreaID, SecondAreaID, FirstAreaHeadID, SecondAreaHeadID As Integer 'first and second area id will store the area id's of the areas the area heads oversee. first and second area head id will store the id's of the area head who will be transferred
    Private FirstAreaHeadYears As Integer ' this will store the amount of years an area head will be based at an area
    Private SecondAreaHeadYears As Integer ' this will store the amount of years an area head will be based at an area
    Private Sub Area_ID_Listbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Area_ID_Listbox.SelectedIndexChanged
        'this sub procedure stores the selected area id that the user selected 
        If Area_ID_Listbox.SelectedItem IsNot Nothing Then
            Selected_Area_ID = Int(Area_ID_Listbox.SelectedItem) ' stores the selected id
        End If
    End Sub

    Private Sub Add_Area_Button_Click(sender As Object, e As EventArgs) Handles Add_Area_Button.Click

        ' when the button is clicked it will carry out some validation checks to make sure invalid data is not being entered. if the fields have been correctly filled it it calls the sub procedure "Create_New_Area" *
        If Area_Name_TextBox.Text = String.Empty Then ' Validation checks on the areas name 
            MsgBox("Please fill in the area's name")
        ElseIf IsNumeric(Area_Name_TextBox.Text) Then
            MsgBox("numbers in names are not permitted")
        ElseIf Len(Area_Name_TextBox.Text) < 4 Or Len(Area_Name_TextBox.Text) > 15 Then
            MsgBox("Please enter a name that is between 4 and 15 characters")


        ElseIf Area_Location_TextBox.Text = String.Empty Then 'Validation checks on the areas location
            MsgBox("Please fill in the areas location")
        ElseIf IsNumeric(Area_Location_TextBox.Text) Then
            MsgBox("numbers are not permitted in a location")

        ElseIf Len(Area_Location_TextBox.Text) < 4 Or Len(Area_Location_TextBox.Text) > 20 Then
            MsgBox("Please enter a location that is between 4 and 20 characters")



        ElseIf Area_Seat_TextBox.Text = String.Empty Then 'Validation checks on the areas seat
            MsgBox("Please fill in the areas seat")
        ElseIf IsNumeric(Area_Seat_TextBox.Text) Then
            MsgBox("numbers are not permitted in a seat")

        ElseIf Len(Area_Seat_TextBox.Text) < 4 Or Len(Area_Seat_TextBox.Text) > 10 Then
            MsgBox("Please enter a seat that is between 4 and 10 characters")

        ElseIf Add_AreaHead_ComboBox.SelectedItem Is Nothing Then 'Validation to check if the area head has been selected 
            MsgBox("Please select an area head")

        Else
            Create_New_Area()
        End If
    End Sub

    Private Sub Create_New_Area()
        ' * this sub procedure stores the new area that has been created in the database *
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            If ValueOfAreaID = -1 Then
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Insert into AREA_TABLE (AREA_NAME, AREA_LOCATION, AREA_SEAT, AREA_HEAD_ID, YEARS_ACTIVE )" & "Values (@AreaName ,@AreaLocation ,@AreaSeat ,@AreaHeadID ,@years  )"
                    .Parameters.AddWithValue("@AreaName", Area_Name_TextBox.Text)
                    .Parameters.AddWithValue("@AreaLocation", Area_Location_TextBox.Text)
                    .Parameters.AddWithValue("@AreaSeat", Area_Seat_TextBox.Text)
                    .Parameters.AddWithValue("@AreaHeadID", Area_Head_ID)
                    .Parameters.AddWithValue("@years ", CurrentYearsActive) ' adding it into the database
                    .ExecuteNonQuery()
                    .CommandText = "Select @@Identity"
                    ValueOfAreaID = .ExecuteScalar
                    Add_Area_Auto_ID_Label.Text = ValueOfAreaID
                    Assign_Leader()
                End With
                Connection.Close() ' closing the connection
                Add_Area_Auto_ID_Label.Text = "Automatically Generated"
                Display_Areas_Table()
            End If

        End If

    End Sub
    Private Sub Assign_Leader()
        '  this sub procedure updates the members table and sets the area head that belongs to the area as a leader 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand

            With SQLCMD
                .Connection = Connection
                .CommandText = "Update MEMBERS_TABLE Set CURRENTLY_LEADING = @Leading " & "Where MEMBER_ID = @ID"
                .Parameters.AddWithValue("@Leading", "Yes") ' setting currently leading to yes 
                .Parameters.AddWithValue("@ID", Area_Head_ID)
                .ExecuteNonQuery()
            End With
        End If
    End Sub
    Private Sub UnAssign_Leader()
        ' this sub procedure uptadates the database to state that the area head is no longer leading the area
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand

            With SQLCMD
                .Connection = Connection
                .CommandText = "Update MEMBERS_TABLE Set CURRENTLY_LEADING = @Leading " & "Where MEMBER_ID = @ID"
                .Parameters.AddWithValue("@Leading", "No") ' setting currently leading to no 
                .Parameters.AddWithValue("@ID", UnAssignedHeadID) ' old area head
                .ExecuteNonQuery()


            End With


        End If

    End Sub
    Private Sub Assign_New_Leader()
        'this sub procedure updates the members table and sets the new area head that belongs to the district as a leader 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Update MEMBERS_TABLE Set CURRENTLY_LEADING = @Leading " & "Where MEMBER_ID = @ID"
                .Parameters.AddWithValue("@Leading", "Yes")
                .Parameters.AddWithValue("@ID", New_Area_Head_ID) ' new area hed
                .ExecuteNonQuery()
            End With
            Connection.Close()
        End If
    End Sub

    Private Sub Clear_Fields()
        ' this sub procedure clears the fields within the form
        Area_ID_Listbox.Items.Clear()
        Area_Name_Listbox.Items.Clear()
        Area_Location_Listbox.Items.Clear()
        Area_Seat_Listbox.Items.Clear()
        Area_Head_Listbox.Items.Clear()
        Area_Name_TextBox.Clear()
        Area_Location_TextBox.Clear()
        Area_Seat_TextBox.Clear()
        New_Area_Name_TextBox.Clear()
        New_Area_Location_TextBox.Clear()
        New_Area_Seat_TextBox.Clear()
        Add_AreaHead_ComboBox.SelectedIndex = -1
        Display_First_Area_Years.Clear()
        Display_Second_Area_Years.Clear()
        New_Area_Head_ComboBox.SelectedIndex = -1
        Second_Area_ComBobox.SelectedIndex = -1
        First_Area_ComBobox.SelectedIndex = -1
        New_Area_Head_ComboBox.Items.Clear()
        First_Area_ComBobox.Items.Clear()
        Second_Area_ComBobox.Items.Clear()
        Add_AreaHead_ComboBox.Items.Clear()
        Add_AreaHead_ComboBox.Items.Clear()
        Display_Area_Name_Textbox.Clear()
        ValueOfAreaID = -1
        Search_Name_Of_Area_Head_TextBox.Clear()
        Search_Name_Of_Area_TextBox.Clear()
        Search_ID_Of_Area_TextBox.Clear()
        Update_Area_Years_NumericUpDown.Value = 0
        Add_Area_Years_NumericUpdown.Value = 0
        First_Area_Head_Name_Textbox.Clear()
        Second_Area_Area_Name_Textbox.Clear()
    End Sub

    Private Sub Display_Areas_Table()
        'this subprocedure displays the details of the areas 
        If DatabaseConnection() Then ' checking for the connection to the databas
            Clear_Fields()
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select MEMBER_NAME,* From MEMBERS_TABLE,AREA_TABLE " & "Where MEMBERS_TABLE.MEMBER_ID = AREA_TABLE.AREA_HEAD_ID"

                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim AreaID, Name, Location, Seat, AreaHead As String ' these variables will hold the data read by the record set 
                    AreaID = DataReader("AREA_ID")
                    Name = DataReader("AREA_NAME")
                    Location = DataReader("AREA_LOCATION")
                    Seat = DataReader("AREA_SEAT")
                    AreaHead = DataReader("MEMBER_NAME")

                    Area_ID_Listbox.Items.Add(AreaID)
                    Area_Name_Listbox.Items.Add(Name)
                    Area_Location_Listbox.Items.Add(Location)
                    Area_Seat_Listbox.Items.Add(Seat)
                    Area_Head_Listbox.Items.Add(AreaHead) ' displaying the data

                End While
                DataReader.Close() ' closing the recordset
            End With
            Connection.Close() ' closing the connection 
        End If


        Load_Areas()
        Load_AreaHeads()
    End Sub
    Private Sub Load_Areas()
        ' this sub procedure loads the names of areas into a combo boxes so that the user can select an area
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select AREA_NAME,AREA_HEAD_ID " & "From MEMBERS_TABLE,AREA_TABLE " & "Where MEMBERS_TABLE.MEMBER_ID = AREA_TABLE.AREA_HEAD_ID "
                .Parameters.AddWithValue("@AreaHead", "Area Head")
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim Name As String ' this will store the names of the areas read by the record set
                    Name = DataReader("AREA_NAME")
                    If Not First_Area_ComBobox.Items.Contains(Name) And Not Second_Area_ComBobox.Items.Contains(Name) Then
                        First_Area_ComBobox.Items.Add(Name)
                        Second_Area_ComBobox.Items.Add(Name) ' adding the names into the combo boxes 
                    End If
                End While
                DataReader.Close()
            End With
            Connection.Close() ' close the connection
        End If
    End Sub

    Private Sub Area_Management_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' this sub procedure calls the procedure that displays all the areas details 
        Display_Areas_Table()
        Delete_Area_Button.Enabled = False
        Update_Area_Button.Enabled = False
    End Sub

    Private Sub Load_AreaHeads()
        'this sub procedure selects the name of the area heads from the database and loads them into the combo box so that it can be selected by the user 

        If DatabaseConnection() Then

            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select MEMBER_NAME From MEMBERS_TABLE Where MEMBER_ROLE = @AreaHead and CURRENTLY_LEADING = @Leading"
                .Parameters.AddWithValue("@AreaHead", "Area Head")
                .Parameters.AddWithValue(" @Leading", "No")
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim AreaHead As String ' this will store the area head name read by the recordset 
                    AreaHead = DataReader("MEMBER_NAME")
                    Add_AreaHead_ComboBox.Items.Add(AreaHead) ' adding it to the combo box
                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If

    End Sub

    Private Sub Load_Area_Name_Button_Click(sender As Object, e As EventArgs) Handles Load_Area_Name_Button.Click
        ' this sub procedure occurs when the user clicks the load button. it calls the function "Load_Area_Name"
        If Area_ID_Listbox.SelectedItem IsNot Nothing Then
            Load_Area_Name()
        End If
    End Sub
    Private Sub Load_Area_Name()
        'this sub procedure selects the name and the ID of the area head that corrosponds to the ID that was selected by the user  
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select AREA_NAME,AREA_HEAD_ID From AREA_TABLE " & "Where AREA_ID = @AreaID"
                .Parameters.AddWithValue("@AreaID", Selected_Area_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim Name As String ' stores the name read by the record set 
                    Name = DataReader("AREA_NAME")
                    UnAssignedHeadID = Int(DataReader("AREA_HEAD_ID"))
                    Display_Area_Name_Textbox.Text = Name ' displaying the name in the textbox
                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If
        Delete_Area_Button.Enabled = True
    End Sub
    Private Sub Delete_Area()
        'this is the actual sub procedure that deletes the area from the database, where the ID is the same as the one selected
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Delete * " & "From AREA_TABLE " & "Where AREA_ID = @DeleteAreaID " ' delete statement
                .Parameters.AddWithValue("@DeleteAreaID", Selected_Area_ID)

                .ExecuteNonQuery()
            End With
            Connection.Close()
            Display_Areas_Table() ' display the info of the areas
            Display_Area_Name_Textbox.Clear() ' clears the textbox
            Delete_Area_Button.Enabled = False
        End If
    End Sub
    Private Sub Assign_New_Area_ID()
        ' this sub procedure selects a new area id that will replace the area that is about to be deleted 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select AREA_ID,NUMBER From AREA_TABLE Where AREA_ID <> @OldAreaID" ' where the area id is not the same as the area id that is about to be deleted
                .Parameters.AddWithValue("@OldAreaID", Selected_Area_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    New_Area_ID = DataReader("AREA_ID") ' storing the new id for the area
                    TemporalAreaNumber = Int(DataReader("NUMBER")) ' storing the number of locals within the district
                End While
                DataReader.Close() ' close the record set

            End With

        End If
    End Sub
    Private Sub Delete_Area_Button_Click(sender As Object, e As EventArgs) Handles Delete_Area_Button.Click
        ' this sub procedure is used to verify if the user is adamant on deleting the area 
        Dim Delete As String 'creating the local variable delete
        Delete = MsgBox("Are you sure you would like to delete this District? The Area's Data will be permenantly deleted, the districts that were assigned to this area will be assigned to another area temporarily(untill you change it)", vbExclamation + vbYesNo + vbDefaultButton2, "Delete Area Confirmation")

        If Delete = vbYes Then ' if yes has been selected 
            Assign_New_Area_ID()
            GetDeletedAreaNumber()
            If DatabaseConnection() Then
                Dim SQLCMD As New OleDbCommand
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Update DISTRICT_TABLE " & " SET AREA_ID = @NewAreaID " & "Where AREA_ID = @OldAreaID" 'update the districts that had this area id with the new area id
                    .Parameters.AddWithValue("@NewDistrict", New_Area_ID) ' new id
                    .Parameters.AddWithValue("@OldAreaID", Selected_Area_ID) ' old id
                    .ExecuteNonQuery()
                    UnAssign_Leader()
                    AddToTemporalArea()
                End With

                Delete_Area() ' call the procedure to delete the area 
            End If
        ElseIf Delete = vbNo Then ' if no is selected 
            Display_Area_Name_Textbox.Clear() ' clear the textbox
            Delete_Area_Button.Enabled = False
        End If
    End Sub
    Private Sub AddToTemporalArea()
        'this sub procedure adds the former number of districts in the area that is being deleted and adds it to the temporal area
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Update AREA_TABLE SET [NUMBER] = @NewNumber " & "Where AREA_ID = @NewAreaID"
                .Parameters.AddWithValue("@NewNumber", Int(TemporalAreaNumber + RemoveNumberInAreas)) ' adding the old area districts to the new area locals
                .Parameters.AddWithValue("@NewLocalID", New_Area_ID)
                .ExecuteNonQuery()

            End With
        End If
    End Sub

    Private Sub GetDeletedAreaNumber()
        'this sub procedure stores the number of districts the deleted area had
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            Dim a As Integer = 0
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select AREA_ID,NUMBER From AREA_TABLE Where AREA_ID = @OldAreaID"
                .Parameters.AddWithValue("@OldDistrictID", Selected_Area_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    RemoveNumberInAreas = Int(DataReader("NUMBER")) '' storing the number

                End While
                DataReader.Close() ' closing the record set

            End With

        End If
    End Sub
    Private Sub Load_New_AreaHeads()
        'this sub procedure loads the names of Area heads into a combo box so the user can select an area head
        If DatabaseConnection() Then

            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select MEMBER_NAME From MEMBERS_TABLE Where MEMBER_ROLE = @AreaHead and CURRENTLY_LEADING = @Leading"
                .Parameters.AddWithValue("@AreaHead", "Area Head")
                .Parameters.AddWithValue("@Leading", "No")
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim AreaHead As String ' this will store the name read by the record set
                    AreaHead = DataReader("MEMBER_NAME")
                    If Not New_Area_Head_ComboBox.Items.Contains(AreaHead) Then
                        New_Area_Head_ComboBox.Items.Add(AreaHead) ' this will store the name read by the record set
                    End If
                End While
                DataReader.Close()
            End With
            Connection.Close() ' close the connection
        End If
    End Sub

    Private Sub Load_Area_Details_Button_Click(sender As Object, e As EventArgs) Handles Load_Area_Details_Button.Click
        'when this button is clicked it calls these sub procedures
        If Area_ID_Listbox.SelectedItem IsNot Nothing Then
            New_Area_Head_ComboBox.Items.Clear()
            Load_Area_Details()
            Load_New_AreaHeads()
        End If
    End Sub
    Private Sub Load_Area_Details()
        'this sub procedure loads the details of the selected area into the text boxes so its information can be modifyied where the area's ID is the same as the one that was selected by the user
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select AREA_NAME,AREA_LOCATION,AREA_SEAT,AREA_HEAD_ID " & "From AREA_TABLE " & "Where AREA_ID = @AreaID "

                .Parameters.AddWithValue("@AreaID", Selected_Area_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim Name, Location, Seat As String ' this will store the values that will be read by the record set 
                    Name = DataReader("AREA_NAME")
                    Location = DataReader("AREA_LOCATION")
                    Seat = DataReader("AREA_SEAT")
                    UnAssignedHeadID = Int(DataReader("AREA_HEAD_ID")) ' storing the current area heads id
                    New_Area_Name_TextBox.Text = Name
                    New_Area_Location_TextBox.Text = Location
                    New_Area_Seat_TextBox.Text = Seat ' displaying the data
                End While
                DataReader.Close()
            End With
            UnAssign_Leader()
            Connection.Close() ' close the connection
            Update_Area_Button.Enabled = True
        End If
    End Sub

    Private Sub Update_Area_Button_Click(sender As Object, e As EventArgs) Handles Update_Area_Button.Click
        ' this subprocedure is called when the update button is clicked. this carries out validation checks on the data the user entered
        If New_Area_Name_TextBox.Text = String.Empty Then ' Validation checks on the areas name 
            MsgBox("Please fill in the area's name")
        ElseIf IsNumeric(New_Area_Name_TextBox.Text) Then
            MsgBox("numbers in names are not permitted")
        ElseIf Len(New_Area_Name_TextBox.Text) < 4 Or Len(New_Area_Name_TextBox.Text) > 15 Then
            MsgBox("Please enter a name that is between 4 and 15 characters")


        ElseIf New_Area_Location_TextBox.Text = String.Empty Then 'Validation checks on the areas location
            MsgBox("Please fill in the areas location")
        ElseIf IsNumeric(New_Area_Location_TextBox.Text) Then
            MsgBox("numbers are not permitted in a location")

        ElseIf Len(New_Area_Location_TextBox.Text) < 4 Or Len(New_Area_Location_TextBox.Text) > 20 Then
            MsgBox("Please enter a location that is between 4 and 20 characters")



        ElseIf New_Area_Seat_TextBox.Text = String.Empty Then 'Validation checks on the areas seat
            MsgBox("Please fill in the areas seat")
        ElseIf IsNumeric(New_Area_Seat_TextBox.Text) Then
            MsgBox("numbers are not permitted in a seat")

        ElseIf Len(New_Area_Seat_TextBox.Text) < 4 Or Len(New_Area_Seat_TextBox.Text) > 10 Then
            MsgBox("Please enter a seat that is between 4 and 10 characters")

        ElseIf New_Area_Head_ComboBox.SelectedItem Is Nothing Then 'Validation to check if the area head has been selected 
            MsgBox("Please select an area head")

        Else
            Update_Member_Details()
        End If
    End Sub

    Private Sub Update_Member_Details()
        'this sub procedure updates the area that corresponds to the area ID that was selected by the user
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Update AREA_TABLE " & "Set AREA_NAME = @AreaName, " & "AREA_LOCATION = @Location, " & "AREA_SEAT = @Seat, " & "AREA_HEAD_ID = @AreaHeadID, " & "YEARS_ACTIVE = @years " & "Where AREA_ID = @AreaID "

                .Parameters.AddWithValue("@AreaName", New_Area_Name_TextBox.Text)
                .Parameters.AddWithValue("@Location", New_Area_Location_TextBox.Text)
                .Parameters.AddWithValue("@Seat", New_Area_Seat_TextBox.Text)
                .Parameters.AddWithValue("@AreaHeadID ", New_Area_Head_ID)
                .Parameters.AddWithValue("@years", NewYearsActive)
                .Parameters.AddWithValue("@AreaID", Selected_Area_ID)

                .ExecuteNonQuery()
                'updates the areas details 
                Update_Area_Button.Enabled = False

            End With

            Connection.Close()
            Assign_New_Leader()
            Display_Areas_Table()
        End If
    End Sub

    Private Sub New_Area_Head_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles New_Area_Head_ComboBox.SelectedIndexChanged
        ' this sub procedure selects the new area head id that corrosponds to the area head the user selected 
        Dim AreaHeadName As String '  holds the name selected from the combo box by the user 
        If New_Area_Head_ComboBox.SelectedItem IsNot Nothing Then
            AreaHeadName = New_Area_Head_ComboBox.SelectedItem.ToString ' stores the name
            If DatabaseConnection() Then
                Dim SQLCMD As New OleDbCommand
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Select MEMBER_ID From MEMBERS_TABLE " & "Where MEMBERS_TABLE.MEMBER_NAME = @AreaHeadName"
                    .Parameters.AddWithValue("@AreaHeadName", AreaHeadName)
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read
                        New_Area_Head_ID = Int(DataReader("MEMBER_ID")) ' storing the new area head id

                    End While
                    DataReader.Close()
                End With
                Connection.Close()
            End If

        End If
    End Sub
    Private Sub Search_Area_Name_Button_Click(sender As Object, e As EventArgs) Handles Search_Area_Name_Button.Click
        'this sub procedure is called when the search by name button is clicked. it carries out validation checks on the name entered by the use
        Dim Area As String
        If Search_Name_Of_Area_TextBox.Text = String.Empty Then ' Validation checks on the areas name 
            MsgBox("Please fill in the area's name")
        ElseIf IsNumeric(Search_Name_Of_Area_TextBox.Text) Then
            MsgBox("numbers in names are not permitted")
        Else
            Area = Search_Name_Of_Area_TextBox.Text
            Search_Area_Name(Area) 'calls the sub procedure 
        End If
    End Sub


    Private Sub Search_Area_Name(AreaName As String) ' the passed in parameter 
        ' this sub procedure searches the database for an area that matches the name entered or a similar name
        If DatabaseConnection() Then
            Clear_Fields()
            Dim SQLcmd As New OleDbCommand
            With SQLcmd
                .Connection = Connection
                .CommandText = "Select *,MEMBER_NAME " & "From AREA_TABLE,MEMBERS_TABLE " & "Where AREA_TABLE.AREA_HEAD_ID = MEMBERS_TABLE.MEMBER_ID and AREA_NAME Like @NameArea " ' if there is a similar name to the area 

                .Parameters.AddWithValue("@NameArea", "%" & AreaName & "%")
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim AreaID, Name, Location, Seat, AreaHead As String ' these variables will store the values read by the record set
                    AreaID = DataReader("AREA_ID")
                    Name = DataReader("AREA_NAME")
                    Location = DataReader("AREA_LOCATION")
                    Seat = DataReader("AREA_SEAT")
                    AreaHead = DataReader("MEMBER_NAME")

                    Area_ID_Listbox.Items.Add(AreaID)
                    Area_Name_Listbox.Items.Add(Name)
                    Area_Location_Listbox.Items.Add(Location)
                    Area_Seat_Listbox.Items.Add(Seat)
                    Area_Head_Listbox.Items.Add(AreaHead) ' display the records

                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If
    End Sub

    Private Sub Search_Area_ID_Button_Click(sender As Object, e As EventArgs) Handles Search_Area_ID_Button.Click
        'this sub procedure is called when the search by ID button is clicked. it carries out validation checks on the ID entered by the user
        Dim ID As Integer
        If Search_ID_Of_Area_TextBox.Text = String.Empty Then ' Validation checks on the areas id
            MsgBox("Please fill in the area's name")
        ElseIf Not IsNumeric(Search_ID_Of_Area_TextBox.Text) Then
            MsgBox("strings are not permitted")
        Else
            ID = Int(Search_ID_Of_Area_TextBox.Text)
            Search_Area_ID(ID) 'calls the sub procedure 
        End If
    End Sub

    Private Sub Search_Area_ID(Area_ID As Integer) ' the passed parameter 
        'this sub procedure searches the database for the record that corresponds to the id that the user typed in the text box
        If DatabaseConnection() Then
            Clear_Fields()
            Dim SQLcmd As New OleDbCommand
            With SQLcmd
                .Connection = Connection
                .CommandText = "Select *,MEMBER_NAME " & "From AREA_TABLE,MEMBERS_TABLE " & "Where AREA_TABLE.AREA_HEAD_ID = MEMBERS_TABLE.MEMBER_ID and AREA_ID = @AreaID " ' and LOCAL_TABLE.LOCAL_ID = MEMBERS_TABLE.LOCAL.ID " 'This is very similar to the StudentID Select query but this time we’re searching using the Surname field and the recordset returned can contain multiple records

                .Parameters.AddWithValue("@AreaID", Area_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim AreaID, Name, Location, Seat, AreaHead As String 'stores the info read by the recordset
                    AreaID = DataReader("AREA_ID")
                    Name = DataReader("AREA_NAME")
                    Location = DataReader("AREA_LOCATION")
                    Seat = DataReader("AREA_SEAT")
                    AreaHead = DataReader("MEMBER_NAME")

                    Area_ID_Listbox.Items.Add(AreaID)
                    Area_Name_Listbox.Items.Add(Name)
                    Area_Location_Listbox.Items.Add(Location)
                    Area_Seat_Listbox.Items.Add(Seat)
                    Area_Head_Listbox.Items.Add(AreaHead) ' display the details 

                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If
    End Sub

    Private Sub Add_AreaHead_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Add_AreaHead_ComboBox.SelectedIndexChanged
        ' this  sub procedure stores the area head id of the area head that is chosen
        Dim AreaHeadName As String ' this will store the name that was selected from the combo box 

        If Add_AreaHead_ComboBox.SelectedItem IsNot Nothing Then
            AreaHeadName = Add_AreaHead_ComboBox.SelectedItem.ToString ' storing the name
            If DatabaseConnection() Then
                Dim SQLCMD As New OleDbCommand
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Select MEMBER_ID From MEMBERS_TABLE " & "Where MEMBERS_TABLE.MEMBER_NAME = @AreaHeadName"
                    .Parameters.AddWithValue("@AreaHeadName", AreaHeadName)
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read
                        Area_Head_ID = Int(DataReader("MEMBER_ID")) ' storing the area head id
                    End While
                    DataReader.Close()
                End With
                Connection.Close()
            End If

        End If
    End Sub


    Private Sub Search_Area_Head_Button_Click(sender As Object, e As EventArgs) Handles Search_Area_Head_Button.Click
        Dim HeadName As String
        If Search_Name_Of_Area_Head_TextBox.Text = String.Empty Then ' Validation checks on the area heads name 
            MsgBox("Please fill in the area head's name")
        ElseIf IsNumeric(Search_Name_Of_Area_Head_TextBox.Text) Then
            MsgBox("numbers in names are not permitted")
        Else
            HeadName = Search_Name_Of_Area_Head_TextBox.Text
            Search_Area_Head_Name(HeadName) 'calls the sub procedure 
        End If
        
    End Sub
    Private Sub Search_Area_Head_Name(AreaHeadsName As String)
        'this sub procedure searches the database for the record that corresponds to the name that the user typed in the text box
        If DatabaseConnection() Then
            Clear_Fields()
            Dim SQLcmd As New OleDbCommand
            With SQLcmd
                .Connection = Connection
                .CommandText = "Select *,MEMBER_NAME " & "From AREA_TABLE,MEMBERS_TABLE " & "Where AREA_TABLE.AREA_HEAD_ID = MEMBERS_TABLE.MEMBER_ID and MEMBER_NAME Like @AreaHead " ' and LOCAL_TABLE.LOCAL_ID = MEMBERS_TABLE.LOCAL.ID " 'This is very similar to the StudentID Select query but this time we’re searching using the Surname field and the recordset returned can contain multiple records

                .Parameters.AddWithValue("@AreaHead", "%" & AreaHeadsName & "%") ' looks for a similar name
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim AreaID, Name, Location, Seat, AreaHead As String
                    AreaID = DataReader("AREA_ID")
                    Name = DataReader("AREA_NAME")
                    Location = DataReader("AREA_LOCATION")
                    Seat = DataReader("AREA_SEAT")
                    AreaHead = DataReader("MEMBER_NAME")

                    Area_ID_Listbox.Items.Add(AreaID)
                    Area_Name_Listbox.Items.Add(Name)
                    Area_Location_Listbox.Items.Add(Location)
                    Area_Seat_Listbox.Items.Add(Seat)
                    Area_Head_Listbox.Items.Add(AreaHead) ' displays the record

                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If
    End Sub

    Private Sub First_Area_ComBobox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles First_Area_ComBobox.SelectedIndexChanged
        ' this sub selects the area id and the amount of years the selected area has 
        Dim AreaName As String ' stores the name chosen by the user 
        If First_Area_ComBobox.SelectedItem IsNot Nothing Then
            AreaName = First_Area_ComBobox.SelectedItem.ToString ' storing the name selected
            If DatabaseConnection() Then
                Dim SQLCMD As New OleDbCommand
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Select AREA_ID,YEARS_ACTIVE From AREA_TABLE " & "Where AREA_TABLE.AREA_NAME = @AreaName"
                    .Parameters.AddWithValue("@AreaName", AreaName)
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read
                        SecondAreaID = Int(DataReader("AREA_ID"))
                        SecondAreaHeadYears = Int(DataReader("YEARS_ACTIVE")) ' storing the data 
                    End While
                    DataReader.Close()
                End With
                Connection.Close()
                Load_Second_Area_Head()
            End If

        End If

    End Sub

    Private Sub Second_Area_ComBobox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Second_Area_ComBobox.SelectedIndexChanged
        ' this sub selects the area id and the amount of years the selected area has 
        Dim AreaName As String ' stores the name chosen by the user 
        If Second_Area_ComBobox.SelectedItem IsNot Nothing Then
            AreaName = Second_Area_ComBobox.SelectedItem.ToString ' storing the name selected

            If DatabaseConnection() Then
                Dim SQLCMD As New OleDbCommand
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Select AREA_ID,YEARS_ACTIVE From AREA_TABLE " & "Where AREA_TABLE.AREA_NAME = @AreaName"
                    .Parameters.AddWithValue("@AreaName", AreaName)
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read
                        FirstAreaID = Int(DataReader("AREA_ID"))
                        FirstAreaHeadYears = Int(DataReader("YEARS_ACTIVE")) ' storing the data 
                    End While
                    DataReader.Close()
                End With
                Connection.Close()
                Load_First_Area_Head()
            End If
        End If
    End Sub
    Private Sub Load_Second_Area_Head()
        ' this sub displays the details of the area head that is about to be transfered
        If DatabaseConnection() Then
            Dim SQLcmd As New OleDbCommand
            With SQLcmd
                .Connection = Connection
                .CommandText = "Select AREA_HEAD_ID,AREA_ID,MEMBER_ID,MEMBER_NAME,YEARS_ACTIVE " & "From AREA_TABLE,MEMBERS_TABLE " & "Where AREA_TABLE.AREA_HEAD_ID = MEMBERS_TABLE.MEMBER_ID and AREA_ID = @AreaID " ' and LOCAL_TABLE.LOCAL_ID = MEMBERS_TABLE.LOCAL.ID " 

                .Parameters.AddWithValue("@AreaID", SecondAreaID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim AreaHead, years As String ' this will store the data that is read
                    SecondAreaHeadID = Int(DataReader("AREA_HEAD_ID"))
                    AreaHead = DataReader("MEMBER_NAME")
                    years = DataReader("YEARS_ACTIVE")
                    First_Area_Head_Name_Textbox.Text = AreaHead
                    Display_First_Area_Years.Text = "years active: " & years ' displaying the data
                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If
    End Sub
    Private Sub Load_First_Area_Head()
        ' this sub displays the details of the area head that is about to be transfered
        If DatabaseConnection() Then
            Dim SQLcmd As New OleDbCommand
            With SQLcmd
                .Connection = Connection
                .CommandText = "Select AREA_HEAD_ID,AREA_ID,MEMBER_ID,MEMBER_NAME,YEARS_ACTIVE " & "From AREA_TABLE,MEMBERS_TABLE " & "Where AREA_TABLE.AREA_HEAD_ID = MEMBERS_TABLE.MEMBER_ID and AREA_ID = @AreaID " ' and LOCAL_TABLE.LOCAL_ID = MEMBERS_TABLE.LOCAL.ID " 

                .Parameters.AddWithValue("@AreaID", FirstAreaID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim AreaHead, years As String ' this will store the data that is read
                    FirstAreaHeadID = Int(DataReader("AREA_HEAD_ID"))
                    AreaHead = DataReader("MEMBER_NAME")
                    years = DataReader("YEARS_ACTIVE")
                    Second_Area_Area_Name_Textbox.Text = AreaHead
                    Display_Second_Area_Years.Text = "years active: " & years ' displaying the data

                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If
    End Sub
    Private Sub Transfer_Area_Head_Button_Click(sender As Object, e As EventArgs) Handles Transfer_Area_Head_Button.Click
        ' this sub procedure checks if all the conditions are met in order to transfer an area head
        If (FirstAreaHeadYears >= 6 And SecondAreaHeadYears >= 6) And (FirstAreaHeadID <> SecondAreaHeadID) Then
            Transfer_First_Area_Head()
            Transfer_Second_Area_Head()
            Display_Areas_Table()

        ElseIf FirstAreaHeadID = SecondAreaHeadID Then ' if same area heads are chosen 
            MsgBox("Select 2 Different Areas with 2 different Area Heads")
        ElseIf FirstAreaHeadYears < 6 Or SecondAreaHeadYears < 6 Then ' if they havent been htere for 6 or more years
            MsgBox("Area Heads can not be transfered if their time has not exceeded 6 years")
        End If
    End Sub

    Private Sub Transfer_First_Area_Head()
        ' this sub procedure updates the area and sets the area head id to the other area head id that belonged to the other area
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Update AREA_TABLE " & "Set AREA_HEAD_ID = @NewAreaHeadID " & "Where AREA_ID = @AreaID"
                .Parameters.AddWithValue("@NewAreaHeadID", FirstAreaHeadID) ' assigning the other head  id 
                .Parameters.AddWithValue("@AreaID", SecondAreaID)
                .ExecuteNonQuery()
            End With
            Connection.Close()

        End If
    End Sub
    Private Sub Transfer_Second_Area_Head()
        ' this sub procedure updates the area and sets the area head id to the other area head id that belonged to the other area
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Update AREA_TABLE " & "Set AREA_HEAD_ID = @NewAreaHeadID " & "Where AREA_ID = @AreaID"
                .Parameters.AddWithValue("@NewAreaHeadID", SecondAreaHeadID) ' assigning the other head  id 
                .Parameters.AddWithValue("@AreaID", FirstAreaID)
                .ExecuteNonQuery()
            End With
            Connection.Close()
        End If
    End Sub

    Private Sub Area_Reports_Button_Click(sender As Object, e As EventArgs) Handles Area_Reports_Button.Click
        '' thsi sub calls the ranking area form to show
        Area_Ranking_Form.ShowDialog()
    End Sub

    Private Sub Add_Area_Years_NumericUpdown_ValueChanged(sender As Object, e As EventArgs) Handles Add_Area_Years_NumericUpdown.ValueChanged
        '' this sub stores the years from the numeric updown
        CurrentYearsActive = Int(Add_Area_Years_NumericUpdown.Value.ToString)
    End Sub

    Private Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles Update_Area_Years_NumericUpDown.ValueChanged
        'this sub stores the years from the numeric updown
        NewYearsActive = Int(Update_Area_Years_NumericUpDown.Value.ToString)
    End Sub

    Private Sub Add_Area_Clear_Button_Click(sender As Object, e As EventArgs) Handles Add_Area_Clear_Button.Click
        'clearing the fields 
        Area_Name_TextBox.Clear()
        Area_Location_TextBox.Clear()
        Area_Seat_TextBox.Clear()
        Add_AreaHead_ComboBox.SelectedIndex = -1
        Add_Area_Years_NumericUpdown.Value = 0
    End Sub

    Private Sub Update_Area_Clear_Button_Click(sender As Object, e As EventArgs) Handles Update_Area_Clear_Button.Click
        'clearing the fields 
        New_Area_Name_TextBox.Clear()
        New_Area_Location_TextBox.Clear()
        New_Area_Seat_TextBox.Clear()
        New_Area_Head_ComboBox.SelectedIndex = -1
        Update_Area_Years_NumericUpDown.Value = 0
    End Sub

    Private Sub Search_Area_Clear_Button_Click(sender As Object, e As EventArgs) Handles Search_Area_Clear_Button.Click
        'clearing the fields 
        Search_Name_Of_Area_TextBox.Clear()
        Search_ID_Of_Area_TextBox.Clear()
    End Sub

    Private Sub Search_Area_Head_Clear_Button_Click(sender As Object, e As EventArgs) Handles Search_Area_Head_Clear_Button.Click
        'clearing the fields 
        Search_Name_Of_Area_Head_TextBox.Clear()
    End Sub

    
End Class