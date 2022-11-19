Imports System.Data.OleDb
Public Class Local_Management_Form
    Private ValueOfLocalID As Integer = -1 ' this variable will hold the value of the local id that is generated
    Private District_ID, New_District_ID As Integer ' these variables will store the district ids that are assinged to the local 
    Private Presiding_Elder_ID, New_Presiding_Elder_ID As Integer ' these variables will store the ids of the presiding elders  
    Private UnAssignedElderID As Integer ' this variable will hold the ID of the presiding elder that will become unassined( not currently leading a local)
    Private Selected_Local_ID As Integer ' this variable will store the id that is selected from the listbox that dislpays the locals id 
    Private NewLocal_ID As Integer 'this variable will hold the new local ID 
    Private TemporalLocalNumber As Integer ' this will hold the number of members in the new temporal local that the members of the old local will be moved to 
    Private AddNumberInDistricts As Integer ' this will store the number of locals that are currently in a district
    Private AddNumberInNewDistrict As Integer ' this will store the number of locals a district has when a new local will be added to it 
    Private SubtractNumberInDistricts As Integer ' this will store the number of how many locals there is in a district before it is decreased by one 
    Private SubtractNumberInDistrictsID As Integer ' this will store the id of the district that is going to have its local number decrease by one
    Private RemoveNumberInLocals As Integer ' this will store the number in the local that is about to be deleted so it can be added onto a tempral local

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Local_ID_Listbox.SelectedIndexChanged
        ' this sub procedure stores the value of the selected id from the local id listbox
        If Local_ID_Listbox.SelectedItem IsNot Nothing Then ' checks if something has been selected
            Selected_Local_ID = Int(Local_ID_Listbox.SelectedItem) ' stores the id that was selected 
        End If
    End Sub
    Private Sub Add_Local_Click(sender As Object, e As EventArgs) Handles Add_Local_Button.Click
        '* when the button is clicked it will carry out some validation checks to make sure invalid data is not being entered. if the fields have been correctly filled it it calls the sub procedure "Create_New_Local" *
        If Local_Name_TextBox.Text = String.Empty Then ' Validation checks on the locals name 
            MsgBox("Please fill in the local's name")
        ElseIf IsNumeric(Local_Name_TextBox.Text) Then
            MsgBox("numbers in names are not permitted")
        ElseIf Len(Local_Name_TextBox.Text) < 4 Or Len(Local_Name_TextBox.Text) > 15 Then
            MsgBox("Please enter a name that is between 4 and 15 characters")


        ElseIf Local_Location_TextBox.Text = String.Empty Then 'Validation checks on the locals location
            MsgBox("Please fill in the locals location")
        ElseIf IsNumeric(Local_Location_TextBox.Text) Then
            MsgBox("numbers are not permitted in a location")

        ElseIf Len(Local_Location_TextBox.Text) < 4 Or Len(Local_Location_TextBox.Text) > 15 Then
            MsgBox("Please enter a location that is between 4 and 15 characters")


        ElseIf Add_Presiding_Elder_ComboBox.SelectedItem Is Nothing Then 'Validation to check if the presiding elder has been selected 
            MsgBox("Please select a presiding elder")
        ElseIf Add_District_ComboBox.SelectedItem Is Nothing Then 'Validation to check if a district has been chosen 
            MsgBox("Please select a district")
        Else
            Create_New_Local()
            AddToDistrict()
        End If
       
    End Sub

    Private Sub Create_New_Local()
        ' * this sub procedure stores the new local that has been created in the database *
        If DatabaseConnection() Then ' checking if there is a database connection 
            Dim SQLCMD As New OleDbCommand 'Open a connection to the database
            If ValueOfLocalID = -1 Then ' if the value of the local id is currently -1
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Insert into LOCAL_TABLE (LOCAL_NAME, LOCAL_LOCATION, PRESIDING_ELDER_ID, DISTRICT_ID )" & "Values (@LocalName ,@LocalLocation ,@PresidingElder ,@DistrictID  )"
                    .Parameters.AddWithValue("@LocalName", Local_Name_TextBox.Text)
                    .Parameters.AddWithValue("@LocalLocation", Local_Location_TextBox.Text)
                    .Parameters.AddWithValue("@PresidingElder", Presiding_Elder_ID)
                    .Parameters.AddWithValue("@DistrictID", District_ID)

                    .ExecuteNonQuery()
                    .CommandText = "Select @@Identity"
                    ValueOfLocalID = .ExecuteScalar 'This is used to run a query which doesn’t return any result. 
                    Add_Local_Auto_ID_Label.Text = ValueOfLocalID ' storing the ID assigned to the local
                    Assign_Leader()
                End With
                Connection.Close()
                Add_Local_Auto_ID_Label.Text = "Automatically Generated"
                Display_Locals_Table() ' calling the procedure that displays the information of the locals
            End If

        End If

    End Sub
    Private Sub Assign_Leader()
        ' this sub procedure updates the members table and sets the presiding elder that beleongs to the loacal as a leader 
        If DatabaseConnection() Then ' check connection
            Dim SQLCMD As New OleDbCommand

            With SQLCMD
                .Connection = Connection
                .CommandText = "Update MEMBERS_TABLE Set CURRENTLY_LEADING = @Leading " & "Where MEMBER_ID = @ID" ' sql update 
                .Parameters.AddWithValue("@Leading", "Yes") ' sets that they are leading 
                .Parameters.AddWithValue("@ID", Presiding_Elder_ID)
                .ExecuteNonQuery() 'This is used to run a query which doesn’t return any result. 
            End With
        End If
    End Sub
    Private Sub AddToDistrict()
        ' this sub procedure adds onto the district that the local belongs to by one
        If DatabaseConnection() Then ' check the connection 
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Update DISTRICT_TABLE Set [NUMBER] = @Num " & "Where DISTRICT_ID = @ID" ' update the number in the district table 

                .Parameters.AddWithValue("@Num", AddNumberInDistricts + 1) ' increment it by one
                ' MsgBox(Local_ID)
                .Parameters.AddWithValue("@ID", District_ID)

                .ExecuteNonQuery() 'This is used to run a query which doesn’t return any result. 
            End With
        End If



    End Sub

    Private Sub AddToTemporalLocal()
        ' this sub procedure adds the number of members who were in the local that is about to be deleted to a temporal local 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Update LOCAL_TABLE SET [NUMBER] = @NewNumber " & "Where LOCAL_ID = @NewLocalID" ' updating the local
                .Parameters.AddWithValue("@NewNumber", Int(TemporalLocalNumber + RemoveNumberInLocals)) ' adding the tempral locals number and the number that the deleted local had 
                .Parameters.AddWithValue("@NewLocalID", NewLocal_ID) ' the temporal local 
                .ExecuteNonQuery() 'This is used to run a query which doesn’t return any result. 

            End With
        End If



    End Sub

    Private Sub GetNumberToDeleteFromDistrict()
        'this sub procedure selects the district id and the number the district has when the local is going to be deleted/ not assined to the district
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select DISTRICT_TABLE.NUMBER,DISTRICT_TABLE.DISTRICT_ID,LOCAL_ID From DISTRICT_TABLE,LOCAL_TABLE " & "Where LOCAL_TABLE.DISTRICT_ID = DISTRICT_TABLE.DISTRICT_ID and LOCAL_TABLE.LOCAL_ID = @ID" ' selecting the district id and the number
                .Parameters.AddWithValue("@ID", Selected_Local_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader() 'This is used to run a select query which may return one or more records. It places those records into a recordset variable
                While DataReader.Read
                    SubtractNumberInDistricts = Int(DataReader("NUMBER"))
                    SubtractNumberInDistrictsID = Int(DataReader("DISTRICT_ID")) ' storing the data read by the recordset
                End While
                DataReader.Close() ' closing the recordset 
            End With
            SubtractFromDistrict() ' calling the sub procedure

        End If
    End Sub
    Private Sub SubtractFromDistrict()
        ' this sub procedure subtracts 1 from the district that the local used to belong to 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Update DISTRICT_TABLE Set [NUMBER] = @Num " & "Where DISTRICT_ID = @ID"

                .Parameters.AddWithValue("@Num", SubtractNumberInDistricts - 1) ' decrease the number by 1
                .Parameters.AddWithValue("@ID", SubtractNumberInDistrictsID) ' where the id is equaivellant to the id stored 

                .ExecuteNonQuery()
            End With
        End If

    End Sub
    Private Sub GetNumberFromDistrict()
        ' this sub procedure  is used to retrive the number of locals that are currently in a particular District
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select NUMBER From DISTRICT_TABLE " & "Where DISTRICT_TABLE.DISTRICT_ID = @DistrictID" ' selecting the number
                .Parameters.AddWithValue("@DistrictID", New_District_ID) ' where the id is equal to  the new district id 
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read

                    AddNumberInNewDistrict = Int(DataReader("NUMBER")) ' storing the number that has been read from the recordset

                End While
                DataReader.Close() ' closing the recordset
            End With
            AddNumberToNewDistrict() 'calling a sub procedure 

        End If
    End Sub
    Private Sub UnAssign_Leader()
        ' this sub procedure uptadates the database to state that the presiding elder is no longer leading the local
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand

            With SQLCMD
                .Connection = Connection
                .CommandText = "Update MEMBERS_TABLE Set CURRENTLY_LEADING  = @Leading " & "Where MEMBER_ID = @ID"
                .Parameters.AddWithValue("@Leading", "No") 'sets the presiding elders leading status to no 
                .Parameters.AddWithValue("@ID", UnAssignedElderID) ' where the id is equal to the one stored 
                .ExecuteNonQuery()


            End With

        End If

    End Sub
    Private Sub AddNumberToNewDistrict()
        'this sub procedure increments the district number that the local has been newly assigned to by 1 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Update DISTRICT_TABLE Set NUMBER = @Num " & "Where DISTRICT_ID = @ID"

                .Parameters.AddWithValue("@Num", AddNumberInNewDistrict + 1) ' adding 1 to the number of locals in the district 
                ' MsgBox(Local_ID)
                .Parameters.AddWithValue("@ID", New_District_ID) ' where the district id is equal to the new district id 

                .ExecuteNonQuery()
            End With
        End If



    End Sub
    Private Sub Assign_New_Leader()
        ' this sub procedure uptadates the database to state that the presiding elder is currently leading the local
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand

            With SQLCMD
                .Connection = Connection
                .CommandText = "Update MEMBERS_TABLE Set CURRENTLY_LEADING  = @Leading " & "Where MEMBER_ID = @ID"
                .Parameters.AddWithValue("@Leading", "Yes") ' setting the presiding elders leading status to yes 
                .Parameters.AddWithValue("@ID", New_Presiding_Elder_ID) ' where the id is the same as the new presiding elder
                .ExecuteNonQuery()
            End With
        End If

    End Sub
    Private Sub Clear_Fields()
        ' this sub procedure clears the text boxes and combo boxes
        Local_ID_Listbox.Items.Clear()
        Search_For_Local_Name_TextBox.Clear()
        Search_For_Local_ID_TextBox.Clear()
        Display_Local_Name_Textbox.Clear()
        Local_Name_Listbox.Items.Clear()
        Local_Location_Listbox.Items.Clear()
        Local_Elders_Listbox.Items.Clear()
        Local_District_Listbox.Items.Clear()
        Local_Area_Listbox.Items.Clear()
        Local_Name_TextBox.Clear()
        Local_Location_TextBox.Clear()
        New_Local_Name_TextBox.Clear()
        New_Local_Location_TextBox.Clear()
        Add_District_ComboBox.SelectedIndex = -1
        Add_Presiding_Elder_ComboBox.SelectedIndex = -1
        New_Presiding_Elder_ComboBox.SelectedIndex = -1
        New_Local_District_ComboBox.SelectedIndex = -1
        New_Presiding_Elder_ComboBox.Items.Clear()
        New_Local_District_ComboBox.Items.Clear()
        Add_District_ComboBox.Items.Clear()
        Add_Presiding_Elder_ComboBox.Items.Clear()
        ValueOfLocalID = -1

    End Sub

    Private Sub Display_Locals_Table()
        'this subprocedure displays the details of the locals and the area and district they have been assigned to 
        If DatabaseConnection() Then ' checking for the connection to the database
            Clear_Fields() ' calling the subprocedure "Clear_Fields"
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select MEMBER_NAME,DISTRICT_NAME,AREA_NAME,* " & "From MEMBERS_TABLE,DISTRICT_TABLE,AREA_TABLE,LOCAL_TABLE " & " Where MEMBERS_TABLE.MEMBER_ID = LOCAL_TABLE.PRESIDING_ELDER_ID and DISTRICT_TABLE.DISTRICT_ID = LOCAL_TABLE.DISTRICT_ID and AREA_TABLE.AREA_ID = DISTRICT_TABLE.AREA_ID  "
                Dim DataReader As OleDbDataReader = .ExecuteReader() ' 'This is used to run a select query which may return one or more records. It places those records into a recordset variable
                While DataReader.Read
                    Dim LocalID, LocalName, Location, ElderName, District, Area As String ' these local varaibles will store the data read by the record set 
                    LocalID = DataReader("LOCAL_ID")
                    LocalName = DataReader("LOCAL_NAME")
                    Location = DataReader("LOCAL_LOCATION")
                    ElderName = DataReader("MEMBER_NAME")
                    District = DataReader("DISTRICT_NAME")
                    Area = DataReader("AREA_NAME")
                    Local_ID_Listbox.Items.Add(LocalID) ' displaying the data that has been read
                    Local_Name_Listbox.Items.Add(LocalName)
                    Local_Location_Listbox.Items.Add(Location)
                    Local_Elders_Listbox.Items.Add(ElderName)
                    Local_District_Listbox.Items.Add(District)
                    Local_Area_Listbox.Items.Add(Area)
                End While
                DataReader.Close() ' closing the recordset
            End With
            Connection.Close() ' closing the connection
            Load_Districts()
            Load_Presiding_Elders() 'calling the sub procedures 
        End If
    End Sub

    Private Sub Load_Presiding_Elders()
        ' this sub procedure is used to load the presiding elder names into the combo box so it can be selected by the user
        If DatabaseConnection() Then

            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select MEMBER_NAME From MEMBERS_TABLE Where MEMBER_ROLE = @PresidingElder and CURRENTLY_LEADING   = @Leading"
                .Parameters.AddWithValue("@PresidingElder", "Presiding Elder")
                .Parameters.AddWithValue(" @Leading", "No")
                Dim DataReader As OleDbDataReader = .ExecuteReader() 'This is used to run a select query which may return one or more records. It places those records into a recordset variable
                While DataReader.Read
                    Dim PresidingElder As String ' creating a local variable to store the names read by the recordset
                    PresidingElder = DataReader("MEMBER_NAME")
                    If Not Add_Presiding_Elder_ComboBox.Items.Contains(PresidingElder) Then ' if its not already in the combo box
                        Add_Presiding_Elder_ComboBox.Items.Add(PresidingElder) 'adding it to the combo box
                    End If
                End While
                DataReader.Close()
            End With

            Connection.Close() ' closing the connection 
        End If


    End Sub

    Private Sub Add_District_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Add_District_ComboBox.SelectedIndexChanged
        'the subprocedure selects the district id and the number of locals that are within the district that is selected from the combobox
        If Add_District_ComboBox.SelectedItem IsNot Nothing Then ' checking if an item has been selected or not
            Dim DistrictName As String ' creating a local variable called  DistrictName. this will hold the name of the local that the user has selected from the combobox
            If Add_District_ComboBox.SelectedItem IsNot Nothing Then
                DistrictName = Add_District_ComboBox.SelectedItem.ToString
                If DatabaseConnection() Then
                    Dim SQLCMD As New OleDbCommand
                    With SQLCMD
                        .Connection = Connection
                        .CommandText = "Select DISTRICT_ID,NUMBER From DISTRICT_TABLE " & "Where DISTRICT_TABLE.DISTRICT_NAME = @DistrictName" 'selecting the district id and the number of locals that corrosponds the district that was selected 
                        .Parameters.AddWithValue("@DistrictName", DistrictName)
                        Dim DataReader As OleDbDataReader = .ExecuteReader()
                        While DataReader.Read
                            District_ID = Int(DataReader("DISTRICT_ID"))
                            AddNumberInDistricts = Int(DataReader("NUMBER")) ' storing the data that the recorset read
                        End While
                        DataReader.Close()
                    End With
                    Connection.Close() ' close the connection
                End If
            End If
        End If
    End Sub

    Private Sub Add_Presiding_Elder_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Add_Presiding_Elder_ComboBox.SelectedIndexChanged
        'this sub procedure selects the presiding elder id that corrosponds to the name of the elder chosen by the user 
        Dim PresidingElderName As String ' this variable will store the name that the user has selected from the combo box 
        If Add_Presiding_Elder_ComboBox.SelectedItem IsNot Nothing Then
            PresidingElderName = Add_Presiding_Elder_ComboBox.SelectedItem.ToString
            If DatabaseConnection() Then
                Dim SQLCMD As New OleDbCommand
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Select MEMBER_ID From MEMBERS_TABLE " & "Where MEMBERS_TABLE.MEMBER_NAME = @PresidingElderName" ' selecting the id
                    .Parameters.AddWithValue("@PresidingElderName", PresidingElderName)
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read
                        Presiding_Elder_ID = Int(DataReader("MEMBER_ID")) ' storing the id
                    End While
                    DataReader.Close() ' closinf the record set
                End With
                Connection.Close() ' close connection
            End If

        End If
    End Sub
    Private Sub Load_Districts()
        'this sub procedure loads the names of districts into the combo box so they can be selected by the user 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select DISTRICT_NAME From DISTRICT_TABLE"

                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim DistrictName As String ' local variable to hold the district name read by the recordset 
                    DistrictName = DataReader("DISTRICT_NAME")
                    Add_District_ComboBox.Items.Add(DistrictName) ' adding the district to the combo box
                End While
                DataReader.Close()
            End With
            Connection.Close() ' closing the connection
        End If
    End Sub

    Private Sub Load_Local_Name_Button_Click(sender As Object, e As EventArgs) Handles Load_Local_Name_Button.Click
        ' this sub procedure occurs when the user clicks the load button. it calls the function "Load_Local_Name"
        If Local_ID_Listbox.SelectedItem IsNot Nothing Then
            Load_Local_Name()
        End If
    End Sub
    Private Sub Load_Local_Name()
        'this sub procedure selcects the name and the ID of the presiding elder that corrosponds to the ID that was selcted by the user  
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select LOCAL_NAME,PRESIDING_ELDER_ID From LOCAL_TABLE " & "Where LOCAL_ID = @LocalID"
                .Parameters.AddWithValue("@LocalID", Selected_Local_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader() 'This is used to run a select query which may return one or more records. It places those records into a recordset variable
                While DataReader.Read
                    Dim LocalName As String ' local variable Localname holds the name of the local the recordset will read from
                    LocalName = DataReader("LOCAL_NAME")
                    UnAssignedElderID = Int(DataReader("PRESIDING_ELDER_ID")) 'storing their id  
                    Display_Local_Name_Textbox.Text = LocalName ' showing the name in the textbox
                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If
        Delete_Local_Button.Enabled = True ' enabling the delete button
    End Sub
    Private Sub Delete_Local()
        'this is the actual sub procedure that deletes the local from the database, where the ID is the same as the one selected
        If DatabaseConnection() Then
            Delete_Assets_For_Local()
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Delete * " & "From LOCAL_TABLE " & "Where LOCAL_ID = @DeleteLocalID " ' the sql command to delete the local
                .Parameters.AddWithValue("@DeleteLocalID", Selected_Local_ID)

                .ExecuteNonQuery() 'This is used to run a query which doesn’t return any result. 
            End With
            Connection.Close()
            Display_Locals_Table()
            Display_Local_Name_Textbox.Clear() ' clears the text box that had the name of the local
            Delete_Local_Button.Enabled = False
        End If
    End Sub
    Private Sub Delete_Assets_For_Local()
        ' this sub procedure deletes all links to the assets in the linking table, so no more assets can be assined to the local that is about to be deleted 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Delete * " & "From LOCAL_ASSET_MAPPING_TABLE " & "Where LOCALS_ID = @DeleteID " 'sql  delete statment
                .Parameters.AddWithValue("@DeleteID", Selected_Local_ID)

                .ExecuteNonQuery()
            End With
        End If
    End Sub
    Private Sub Delete_Local_Button_Click(sender As Object, e As EventArgs) Handles Delete_Local_Button.Click
        Dim Delete As String 'creating the local variable delete
        Delete = MsgBox("Are you sure you would like to delete this Local? The Local's Data will be permenantly deleted, the members and services that were assigned to this local will be assigned to another local temporarily(untill you change it)", vbExclamation + vbYesNo + vbDefaultButton2, "Delete Local Confirmation") 'message box alerts the user this data will be permenatly lost

        If Delete = vbYes Then ' if yess is chosen
            Assign_New_Local_ID()
            GetDeletedLocalNumber() ' calling sub procedures 

            If DatabaseConnection() Then
                Dim SQLCMD As New OleDbCommand
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Update MEMBERS_TABLE " & " SET LOCALS_ID = @NewLocalID " & "Where LOCALS_ID = @OldLocalID" ' updating the local id of all the members that have been assingned to the local that is about to be deleted 
                    .Parameters.AddWithValue("@NewLocalID", NewLocal_ID)
                    .Parameters.AddWithValue("@OldLocalID", Selected_Local_ID)
                    .ExecuteNonQuery()
                    .CommandText = "Update SERVICE_TABLE " & " SET LOCAL_ID = @NewLocalID " & "Where LOCAL_ID = @OldLocalID" ' updating the local id of all the services that have been assingned to the local that is about to be deleted 
                    .Parameters.AddWithValue("@NewLocalID", NewLocal_ID)
                    .Parameters.AddWithValue("@OldLocalID", Selected_Local_ID)
                    .ExecuteNonQuery()
                    UnAssign_Leader()
                    GetNumberToDeleteFromDistrict()
                    AddToTemporalLocal() ' calling sub procedures 
                End With

                Delete_Local() ' calling the function that actually deletes the local
            End If
        ElseIf Delete = vbNo Then ' if no is selected 
            Display_Local_Name_Textbox.Clear() ' clear the textbox
            Delete_Local_Button.Enabled = False
        End If
    End Sub
    Private Sub GetDeletedLocalNumber()
        ' this sub procedure retrives the number of members the local that is about to be deleted has 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select LOCAL_ID,NUMBER From LOCAL_TABLE Where LOCAL_ID = @OldLocalID"
                .Parameters.AddWithValue("@OldLocalID", Selected_Local_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    RemoveNumberInLocals = Int(DataReader("NUMBER")) ' storing the number of members that was in that local

                End While
                DataReader.Close()
            End With

        End If
    End Sub
    Private Sub Assign_New_Local_ID()
        ' this sub procedure selects a new local id that will replace the local that is about to be deleted 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select LOCAL_ID,NUMBER From LOCAL_TABLE Where LOCAL_ID <> @OldLocalID" ' selectin a local id where the local id is not the same as the local id that is about to be deleted 
                .Parameters.AddWithValue("@OldLocalID", Selected_Local_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    NewLocal_ID = DataReader("LOCAL_ID") ' storing the local id 
                    TemporalLocalNumber = Int(DataReader("NUMBER")) ' storing the number of members who were in the local 


                End While
                DataReader.Close() 'closing the recordset 

            End With

        End If
    End Sub
    Private Sub Load_New_Districts()
        'this sub procedure loads the names of the districts into the combo box so it can be selected by the user 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select DISTRICT_NAME From DISTRICT_TABLE" ' selecting the name
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim DistrictName As String ' variable holds the name  read by the recordset 
                    DistrictName = DataReader("DISTRICT_NAME")
                    New_Local_District_ComboBox.Items.Add(DistrictName) ' adding the district name to the combo box 
                End While
                DataReader.Close()
            End With
            Connection.Close() ' close the connection
        End If
    End Sub
    Private Sub Load_New_Presiding_Elders()
        ' this sub procedure loads the names of presiding eders into the combo box so they can be selected by the user
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select MEMBER_NAME From MEMBERS_TABLE Where MEMBER_ROLE = @PresidingElder and CURRENTLY_LEADING = @Leading " ' selecting the name where the role is "Presiding elder" and they are not currently leading 
                .Parameters.AddWithValue("@PresidingElder", "Presiding Elder")
                .Parameters.AddWithValue("@Leading", "No")
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim PresidingElder As String 'stores the name that will be read by the recordset 
                    PresidingElder = DataReader("MEMBER_NAME")
                    If Not New_Presiding_Elder_ComboBox.Items.Contains(PresidingElder) Then ' if the combo box doesnt have the name
                        New_Presiding_Elder_ComboBox.Items.Add(PresidingElder) ' add it to the combo box
                    End If
                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If




    End Sub

    Private Sub Load_Local_Details_Button_Click(sender As Object, e As EventArgs) Handles Load_Local_Details_Button.Click
        ' when this button is clicked it calls these sub procedures 
        If Local_ID_Listbox.SelectedItem IsNot Nothing Then
            New_Local_District_ComboBox.Items.Clear()
            New_Local_District_ComboBox.Items.Clear()
            Load_New_Districts()
            Load_Local_Details()
            Load_New_Presiding_Elders()
            ' calls these sub procedures
        End If
    End Sub
    Private Sub Load_Local_Details()
        'this sub procedure loads the details of the selected local into the text boxes so its information can be modifyied where the local's ID is the same as the one that was selected by the user
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select * " & "From LOCAL_TABLE " & "Where LOCAL_ID = @LocalID "

                .Parameters.AddWithValue("@LocalID", Selected_Local_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim LocalName, Location As String 'these variables are created in order to hold the data that was read by the recordset 
                    LocalName = DataReader("LOCAL_NAME")
                    Location = DataReader("LOCAL_LOCATION")
                    District_ID = DataReader("DISTRICT_ID")
                    UnAssignedElderID = Int(DataReader("PRESIDING_ELDER_ID"))
                    New_Local_Name_TextBox.Text = LocalName
                    New_Local_Location_TextBox.Text = Location ' display the info in the textbox
                End While
                DataReader.Close()
            End With
            UnAssign_Leader()
            Connection.Close()
            Update_Local_Button.Enabled = True ' enabling the update button 
        End If
    End Sub

    Private Sub Update_Local_Button_Click(sender As Object, e As EventArgs) Handles Update_Local_Button.Click
        ' this subprocedure is called when the update button is clicked. this carries out validation checks on the data the user entered
        If New_Local_Name_TextBox.Text = String.Empty Then ' Validation checks on the locals name 
            MsgBox("Please fill in the local's name")
        ElseIf IsNumeric(New_Local_Name_TextBox.Text) Then
            MsgBox("numbers in names are not permitted")
        ElseIf Len(New_Local_Name_TextBox.Text) < 4 Or Len(Local_Name_TextBox.Text) > 15 Then
            MsgBox("Please enter a name that is between 4 and 15 characters")


        ElseIf New_Local_Location_TextBox.Text = String.Empty Then 'Validation checks on the locals location
            MsgBox("Please fill in the locals location")
        ElseIf IsNumeric(New_Local_Location_TextBox.Text) Then
            MsgBox("numbers are not permitted in a location")

        ElseIf Len(New_Local_Location_TextBox.Text) < 4 Or Len(Local_Location_TextBox.Text) > 15 Then
            MsgBox("Please enter a location that is between 4 and 15 characters")


        ElseIf New_Presiding_Elder_ComboBox.SelectedItem Is Nothing Then 'Validation to check if the presiding elder has been selected 
            MsgBox("Please select a presiding elder")
        ElseIf New_Local_District_ComboBox.SelectedItem Is Nothing Then 'Validation to check if a district has been chosen 
            MsgBox("Please select a district")
        Else
            Update_Local_Details()
        End If

    End Sub
    Private Sub UpdateNumbersInDistrict()
        ' 'this sub procedure is used to determine how the district will be updated, whether they will lose a local or not when a new local has been updated
        If DatabaseConnection() Then
            If District_ID <> New_District_ID Then ' if the district id that the local originally had is the same as its new one
                Dim SQLCMD As New OleDbCommand
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Select NUMBER,DISTRICT_ID From DISTRICT_TABLE " & "Where DISTRICT_TABLE.DISTRICT_ID = @DistrictID"
                    .Parameters.AddWithValue("@DistrictID", District_ID)
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read

                        SubtractNumberInDistrictsID = Int(DataReader("DISTRICT_ID"))
                        SubtractNumberInDistricts = Int(DataReader("NUMBER"))
                        ' storing the id and the number currently in the district 
                    End While
                    DataReader.Close()

                End With
                SubtractFromDistrict()
                GetNumberFromDistrict() ' calling the sub procedures 

            End If
        End If

    End Sub
    Private Sub Update_Local_Details()
        'this sub procedure updates the local that corresponds to the local ID that was selected by the user
        If DatabaseConnection() Then
            UpdateNumbersInDistrict() ' calls the sub procedure above 
            Assign_New_Leader()
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Update LOCAL_TABLE " & "Set LOCAL_NAME = @LocalName, " & "LOCAL_LOCATION = @Location, " & "PRESIDING_ELDER_ID = @ElderID, " & "DISTRICT_ID = @DistrictID " & "Where LOCAL_ID = @LocalID "

                .Parameters.AddWithValue("LocalName", New_Local_Name_TextBox.Text)
                .Parameters.AddWithValue("@Location", New_Local_Location_TextBox.Text)
                .Parameters.AddWithValue("ElderID", New_Presiding_Elder_ID)
                .Parameters.AddWithValue("@DistrictID", New_District_ID)
                .Parameters.AddWithValue("@LocalID", Selected_Local_ID)
                .ExecuteNonQuery()
                Update_Local_Button.Enabled = False
                'updates the locals details 

            End With

            Connection.Close() ' close connection
            Display_Locals_Table() ' show local details
        End If
    End Sub

    Private Sub New_Presiding_Elder_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles New_Presiding_Elder_ComboBox.SelectedIndexChanged
        ' this sub procedure  obtains the presiding elder new ID
        Dim PresidingElderName As String ' holds the name selected from the combo box by the user 
        If New_Presiding_Elder_ComboBox.SelectedItem IsNot Nothing Then ' if there has been a name selected 
            PresidingElderName = New_Presiding_Elder_ComboBox.SelectedItem.ToString ' store the name to the variable
            If DatabaseConnection() Then
                Dim SQLCMD As New OleDbCommand
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Select MEMBER_ID From MEMBERS_TABLE " & "Where MEMBERS_TABLE.MEMBER_NAME = @ElderName"
                    .Parameters.AddWithValue("@ElderName", PresidingElderName)
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read
                        New_Presiding_Elder_ID = Int(DataReader("MEMBER_ID")) ' selecting the id of the presiding elder 

                    End While
                    DataReader.Close()
                End With
                Connection.Close() ' close the connection
            End If
        End If
    End Sub

    Private Sub New_Local_District_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles New_Local_District_ComboBox.SelectedIndexChanged
        'this procedure stores the new district ID
        Dim DistrictName As String ' this will store the district name that the user selects
        If New_Local_District_ComboBox.SelectedItem IsNot Nothing Then
            DistrictName = New_Local_District_ComboBox.SelectedItem.ToString ' storing the name
            If DatabaseConnection() Then
                Dim SQLCMD As New OleDbCommand
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Select DISTRICT_ID From DISTRICT_TABLE " & "Where DISTRICT_TABLE.DISTRICT_NAME = @DistrictName"
                    .Parameters.AddWithValue("@DistrictName", DistrictName)
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read
                        New_District_ID = Int(DataReader("DISTRICT_ID")) ' storing the id that matches to the district name

                    End While
                    DataReader.Close()
                End With
                Connection.Close()
            End If

        End If
    End Sub

    Private Sub Search_Local_Name_Button_Click(sender As Object, e As EventArgs) Handles Search_Local_Name_Button.Click
        'this sub procedure is called when the search by name button is clicked. it carries out validation checks on the name entered by the user
        Dim Name As String ' this will store the name that the user types into the search textbox 
        If Search_For_Local_Name_TextBox.Text = String.Empty Then ' validation checks on the name
            MsgBox("Please fill in the locals name")
        ElseIf IsNumeric(Search_For_Local_Name_TextBox.Text) Then
            MsgBox("numbers in names are not permitted")
        Else
            Name = Search_For_Local_Name_TextBox.Text
            Search_Local_Name(Name) ' calls the sub procedure 
        End If
    End Sub


    Private Sub Search_Local_Name(Name As String)
        ' this sub procedure searches the database for a local that matches the name entered or a similar name
        If DatabaseConnection() Then
            Clear_Fields() ' clearing the fields
            Dim SQLcmd As New OleDbCommand
            With SQLcmd
                .Connection = Connection
                .CommandText = "Select MEMBER_NAME,DISTRICT_NAME,AREA_NAME,* " & "From MEMBERS_TABLE,DISTRICT_TABLE,AREA_TABLE,LOCAL_TABLE " & " Where LOCAL_NAME Like @LocalName and MEMBERS_TABLE.MEMBER_ID = LOCAL_TABLE.PRESIDING_ELDER_ID and DISTRICT_TABLE.DISTRICT_ID = LOCAL_TABLE.DISTRICT_ID and AREA_TABLE.AREA_ID = DISTRICT_TABLE.AREA_ID  " 'if there is a similar name
                .Parameters.AddWithValue("@LocalName", "%" & Name & "%")
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim LocalID, LocalName, Location, PresidingElder, District, Area As String '' these variables store the data read by the record set 
                    LocalID = DataReader("LOCAL_ID")
                    LocalName = DataReader("LOCAL_NAME")
                    Location = DataReader("LOCAL_LOCATION")
                    PresidingElder = DataReader("MEMBER_NAME")
                    District = DataReader("DISTRICT_NAME")
                    Area = DataReader("AREA_NAME")
                    Local_ID_Listbox.Items.Add(LocalID)
                    Local_Name_Listbox.Items.Add(LocalName)
                    Local_Location_Listbox.Items.Add(Location)
                    Local_Elders_Listbox.Items.Add(PresidingElder)
                    Local_District_Listbox.Items.Add(District)
                    Local_Area_Listbox.Items.Add(Area)

                End While

            End With
            Connection.Close()
        End If

    End Sub

    Private Sub Search_Local_ID_Button_Click(sender As Object, e As EventArgs) Handles Search_Local_ID_Button.Click
        'this sub procedure is called when the search by ID button is clicked. it carries out validation checks on the ID entered by the user
        Dim ID As Integer 'storing the ID entered by the user
        If Search_For_Local_ID_TextBox.Text = String.Empty Then ' validation checks 
            MsgBox("Please fill in the local's ID")
        ElseIf Not IsNumeric(Search_For_Local_ID_TextBox.Text) Then
            MsgBox("strings are not permitted in an ID")
        Else
            ID = Int(Search_For_Local_ID_TextBox.Text)
            Search_Local_ID(ID) 'passes the parameter through
        End If

    End Sub

    Private Sub Search_Local_ID(Local_ID As Integer) ' the passed parameter 
        'this sub procedure searches the database for the record that corresponds to the id that the user typed in the text box
        If DatabaseConnection() Then
            Clear_Fields()
            Dim SQLcmd As New OleDbCommand
            With SQLcmd
                .Connection = Connection
                .CommandText = "Select MEMBER_NAME,DISTRICT_NAME,AREA_NAME,* " & "From MEMBERS_TABLE,DISTRICT_TABLE,AREA_TABLE,LOCAL_TABLE " & " Where LOCAL_ID = @LocalID and MEMBERS_TABLE.MEMBER_ID = LOCAL_TABLE.PRESIDING_ELDER_ID and DISTRICT_TABLE.DISTRICT_ID = LOCAL_TABLE.DISTRICT_ID and AREA_TABLE.AREA_ID = DISTRICT_TABLE.AREA_ID  "

                .Parameters.AddWithValue("@LocalID", Local_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim LocalID, LocalName, Location, PresidingElder, District, Area As String ''stores the info read by the recordset 
                    LocalID = DataReader("LOCAL_ID")
                    LocalName = DataReader("LOCAL_NAME")
                    Location = DataReader("LOCAL_LOCATION")
                    PresidingElder = DataReader("MEMBER_NAME")
                    District = DataReader("DISTRICT_NAME")
                    Area = DataReader("AREA_NAME")
                    Local_ID_Listbox.Items.Add(LocalID)
                    Local_Name_Listbox.Items.Add(LocalName)
                    Local_Location_Listbox.Items.Add(Location)
                    Local_Elders_Listbox.Items.Add(PresidingElder)
                    Local_District_Listbox.Items.Add(District)
                    Local_Area_Listbox.Items.Add(Area)
                    ' output the records found
                End While

            End With
            Connection.Close()
        End If
    End Sub

    Private Sub Local_Management_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' when the form loads it calls the procedure to display the locals info and disables some buttons
        Display_Locals_Table()
        Delete_Local_Button.Enabled = False
        Update_Local_Button.Enabled = False
    End Sub
    Private Sub Local_Reports_Button_Click(sender As Object, e As EventArgs) Handles Local_Reports_Button.Click
        ' when this button is clicked it opens the local ranking form
        Local_Ranking_Form.ShowDialog()
    End Sub

    Private Sub Add_Local_Clear_Button_Click(sender As Object, e As EventArgs) Handles Add_Local_Clear_Button.Click
        ' clears the user inputs
        Add_District_ComboBox.SelectedIndex = -1
        Add_Presiding_Elder_ComboBox.SelectedIndex = -1
        Local_Name_TextBox.Clear()
        Local_Location_TextBox.Clear()


    End Sub

    Private Sub Search_Local_Clear_Button_Click(sender As Object, e As EventArgs) Handles Search_Local_Clear_Button.Click
        ' clears the user inputs
        Search_For_Local_ID_TextBox.Clear()
        Search_For_Local_Name_TextBox.Clear()
    End Sub

    Private Sub Update_Local_Clear_Button_Click(sender As Object, e As EventArgs) Handles Update_Local_Clear_Button.Click
        ' clears the user inputs
        New_Local_District_ComboBox.SelectedIndex = -1
        New_Presiding_Elder_ComboBox.SelectedIndex = -1
        New_Local_Location_TextBox.Clear()
        New_Local_Name_TextBox.Clear()
    End Sub
End Class