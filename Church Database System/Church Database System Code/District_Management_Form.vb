Imports System.Data.OleDb
Public Class District_Management_Form
    Private ValueOfDistrictID As Integer = -1 ' this variable will hold the value of the district id 
    Private District_Pastor_ID As Integer ' this variable will hold the id of the district pastor 
    Private New_District_Pastor_ID As Integer ' this variable will hold the id of the new district pastor 
    Private Selected_District_ID As Integer ' this variable will store the district id that the user selected 
    Private New_Area_ID, Area_ID As Integer ' area and new area id will hold the area id's that the district is assigned to 
    Private New_District_ID As Integer ' this variable will store the new district id that will be assigned to locals temperaly when the current district is being deleted 
    Private FirstAreaID As Integer ' this variable will store the first area id
    Private SecondAreaID As Integer ' this variable will store the second area id 
    Private TemporalDistrictNumber As Integer ' this variable will hold the number of local that are in the temporal district before another districts number is added onto it 
    Private UnAssignedPastorID As Integer ' this variable will store the id of the district pastor who will become un assinged
    Private AddNumberInAreas As Integer 'this variable will store the number of districts there are in an area 
    Private AddNumberInNewArea As Integer ' this variable stores the number of districts in a new area before its incremented by 1
    Private SubtractNumberInAreas As Integer ' this variable stores the amount of districts in an area before it is decreased by 1
    Private SubtractNumberInAreasID As Integer ' this variable stores the area id of the area that will have a district removed
    Private RemoveNumberInDistricts As Integer ' this variable stores the number of locals a district has before it is deleted 
    Private FirstDistrictPastorYears As Integer ' this variable will store how many years the first district pastor has been at a district
    Private SecondDistrictPastorYears As Integer ' this variable will store how many years the Second district pastor has been at a district
    Private CurrentYearsActive As Integer ' this variable will store the number of years a district pastor has been based at a district
    Private NewYearsActive As Integer 'this variable will store the new number of years a district pastor has been based at a district
    Private FirstDistrictID, SecondDistrictID, FirstDistrictPastorID, SecondDistrictPastorID As Integer 'First and Second district pastor id store the id of the pastors who will be transfered . Second and First district id stores what district id the district pastors are under
    Private Sub District_ID_Listbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles District_ID_Listbox.SelectedIndexChanged
        ' this sub procedure stores the selected district id that the user selected 
        If District_ID_Listbox.SelectedItem IsNot Nothing Then
            Selected_District_ID = Int(District_ID_Listbox.SelectedItem) ' storing the district ID

        End If
    End Sub

    Private Sub Add_District_Button_Click(sender As Object, e As EventArgs) Handles Add_District_Button.Click
        ' when the button is clicked it will carry out some validation checks to make sure invalid data is not being entered. if the fields have been correctly filled it it calls the sub procedure "Create_New_District" *
        If District_Name_TextBox.Text = String.Empty Then ' Validation checks on the districts name 
            MsgBox("Please fill in the district's name")
        ElseIf IsNumeric(District_Name_TextBox.Text) Then
            MsgBox("numbers in names are not permitted")
        ElseIf Len(District_Name_TextBox.Text) < 4 Or Len(District_Name_TextBox.Text) > 15 Then
            MsgBox("Please enter a name that is between 4 and 15 characters")


        ElseIf District_Location_TextBox.Text = String.Empty Then 'Validation checks on the districts location
            MsgBox("Please fill in the districts location")
        ElseIf IsNumeric(District_Location_TextBox.Text) Then
            MsgBox("numbers are not permitted in a location")

        ElseIf Len(District_Location_TextBox.Text) < 4 Or Len(District_Location_TextBox.Text) > 15 Then
            MsgBox("Please enter a location that is between 4 and 15 characters")


        ElseIf Add_District_Pastor_ComboBox.SelectedItem Is Nothing Then 'Validation to check if the district pastor has been selected 
            MsgBox("Please select a district pastor")
        ElseIf Add_Area_ComboBox.SelectedItem Is Nothing Then 'Validation to check if an area has been chosen 
            MsgBox("Please select an area")
        Else
            Create_New_District()
            AddToArea()
        End If
    End Sub

    Private Sub AddToArea()
        ' this procedure increments the area a district is assigned to by 1
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Update AREA_TABLE Set [number] = @Num " & "Where AREA_ID = @ID"

                .Parameters.AddWithValue("@Num", AddNumberInAreas + 1) ' adding 1 onto the current number  
                .Parameters.AddWithValue("@ID", Area_ID)

                .ExecuteNonQuery()
            End With
        End If
    End Sub
    Private Sub Create_New_District()
        ' * this sub procedure stores the new district that has been created in the database *
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            If ValueOfDistrictID = -1 Then
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Insert into DISTRICT_TABLE (DISTRICT_NAME, DISTRICT_LOCATION, DISTRICT_PASTOR_ID, AREA_ID,YEARS_ACTIVE )" & "Values (@DistrictName ,@Location ,@DistrictPastor ,@AreaID,@years )"
                    .Parameters.AddWithValue("@DistrictName", District_Name_TextBox.Text)
                    .Parameters.AddWithValue("@Location ", District_Location_TextBox.Text)
                    .Parameters.AddWithValue("@DistrictPastor", District_Pastor_ID)
                    .Parameters.AddWithValue("@AreaID ", Area_ID)

                    .Parameters.AddWithValue("@years ", CurrentYearsActive) ' adding the data into the database 
                    .ExecuteNonQuery()
                    .CommandText = "Select @@Identity"
                    ValueOfDistrictID = .ExecuteScalar
                    Add_District_Auto_ID_Label.Text = ValueOfDistrictID
                    Assign_Leader()
                End With
                Connection.Close()
                Add_District_Auto_ID_Label.Text = "Automatically Generated"
                Display_Districts_Table()
            End If

        End If

    End Sub
    Private Sub Assign_Leader()
        ' this sub procedure updates the members table and sets the district pastor that belongs to the district as a leader 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand

            With SQLCMD
                .Connection = Connection
                .CommandText = "Update MEMBERS_TABLE Set CURRENTLY_LEADING = @Leading " & "Where MEMBER_ID = @ID"
                .Parameters.AddWithValue("@Leading", "Yes") ' setting currently leading to yes 
                .Parameters.AddWithValue("@ID", District_Pastor_ID)
                .ExecuteNonQuery()
            End With
        End If
    End Sub
    Private Sub Assign_New_Leader()
        'this sub procedure updates the members table and sets the new district pastor that belongs to the district as a leader 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand

            With SQLCMD
                .Connection = Connection
                .CommandText = "Update MEMBERS_TABLE Set CURRENTLY_LEADING = @Leading " & "Where MEMBER_ID = @ID"
                .Parameters.AddWithValue("@Leading", "Yes") ' setting currently leading to yes 
                .Parameters.AddWithValue("@ID", New_District_Pastor_ID)
                .ExecuteNonQuery()
            End With
            Connection.Close() ' closing the connection 
        End If

    End Sub
    Private Sub UnAssign_Leader()
        If DatabaseConnection() Then
            ' this sub procedure uptadates the database to state that the district pastor is no longer leading the district
            Dim SQLCMD As New OleDbCommand

            With SQLCMD
                .Connection = Connection
                .CommandText = "Update MEMBERS_TABLE Set CURRENTLY_LEADING = @Leading " & "Where MEMBER_ID = @ID"
                .Parameters.AddWithValue("@Leading", "No") ' setting currently leading to no 
                .Parameters.AddWithValue("@ID", UnAssignedPastorID)
                .ExecuteNonQuery()
            End With
        End If
    End Sub
    Private Sub Clear_Fields()
        ' this sub procedure clears the fields within the form
        District_ID_Listbox.Items.Clear()
        District_Name_Listbox.Items.Clear()
        District_Location_Listbox.Items.Clear()
        District_Pastor_Listbox.Items.Clear()
        District_Area_Listbox.Items.Clear()
        District_Name_TextBox.Clear()
        District_Location_TextBox.Clear()
        Second_District_Pastor_Name_Textbox.Clear()
        First_District_Pastor_Name_Textbox.Clear()
        New_District_Name_TextBox.Clear()
        New_District_Location_TextBox.Clear()
        New_District_Pastor_ComboBox.SelectedIndex = -1
        Add_Area_ComboBox.SelectedIndex = -1
        Add_District_Pastor_ComboBox.SelectedIndex = -1
        New_District_Area_ComboBox.SelectedIndex = -1
        Second_District_ComBobox.SelectedIndex = -1
        First_District_ComBobox.SelectedIndex = -1
        New_District_Area_ComboBox.Items.Clear()
        New_District_Pastor_ComboBox.Items.Clear()
        First_District_ComBobox.Items.Clear()
        Second_District_ComBobox.Items.Clear()
        Add_Area_ComboBox.Items.Clear()
        Add_District_Pastor_ComboBox.Items.Clear()
        ValueOfDistrictID = -1
        Update_District_Years_NumericUpDown.Value = 0
        District_Years_NumericUpDown.Value = 0
        Search_For_District_ID_TextBox.Clear()
        Search_For_District_Name_TextBox.Clear()
        Search_For_District_Pastor_Name_TextBox.Clear()
        Display_Second_Years_Textbox.Clear()
        Display_First_Years_Textbox.Clear()
        Display_District_Name_Textbox.Clear()
    End Sub

    Private Sub Display_Districts_Table()
        'this subprocedure displays the details of the districts and the area they have been assigned to 
        If DatabaseConnection() Then ' checking for the connection to the database
            Clear_Fields() ' calling the subprocedure "Clear_Fields"
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select MEMBER_NAME,AREA_NAME,* From MEMBERS_TABLE,AREA_TABLE,DISTRICT_TABLE " & "Where MEMBERS_TABLE.MEMBER_ID = DISTRICT_TABLE.DISTRICT_PASTOR_ID and AREA_TABLE.AREA_ID = DISTRICT_TABLE.AREA_ID " ' selecting the districts and their district pastor and area they are under

                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim DistrictID, Name, Location, DistrictPastor, Area As String ' these variables will hold the data read by the record set 
                    DistrictID = DataReader("DISTRICT_ID")
                    Name = DataReader("DISTRICT_NAME")
                    Location = DataReader("DISTRICT_LOCATION")
                    DistrictPastor = DataReader("MEMBER_NAME")
                    Area = DataReader("AREA_NAME")

                    District_ID_Listbox.Items.Add(DistrictID)
                    District_Name_Listbox.Items.Add(Name)
                    District_Location_Listbox.Items.Add(Location)
                    District_Pastor_Listbox.Items.Add(DistrictPastor)
                    District_Area_Listbox.Items.Add(Area) ' displaying the data

                End While
                DataReader.Close() ' close the recorset
            End With
            Connection.Close() ' close the connection
        End If

        Load_Areas()
        Load_District_Pastors()
        Load_Transfer_District() 'call sub procedures

    End Sub

    Private Sub Load_District_Pastors()
        ' this sub procedure loads the names of district pastors into a combo box so that the user can select a district pastor 
        If DatabaseConnection() Then ' checks if there is a connection 
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select MEMBER_NAME,MEMBER_ID From MEMBERS_TABLE Where MEMBER_ROLE = @DistrictPastor and CURRENTLY_LEADING = @Leading"
                .Parameters.AddWithValue(" @DistrictPastor", "District Pastor")
                .Parameters.AddWithValue(" @Leading", "No") ' select a district pastor who is not currently leading another district 

                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim DistrictPastor As String

                    DistrictPastor = DataReader("MEMBER_NAME") ' holds the name that the record set will read from 
                    If Not Add_District_Pastor_ComboBox.Items.Contains(DistrictPastor) Then
                        Add_District_Pastor_ComboBox.Items.Add(DistrictPastor) ' adds it into the combo box
                    End If
                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If

    End Sub
    Private Sub Load_Transfer_District()
        ' this sub procedure loads the names of districts into a combo boxes so that the user can select a district
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select DISTRICT_NAME,DISTRICT_PASTOR_ID " & "From MEMBERS_TABLE,DISTRICT_TABLE " & "Where MEMBERS_TABLE.MEMBER_ID = DISTRICT_TABLE.DISTRICT_PASTOR_ID "
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim Name As String ' stores the name read by the record set 
                    Name = DataReader("DISTRICT_NAME")
                    If Not First_District_ComBobox.Items.Contains(Name) And Not Second_District_ComBobox.Items.Contains(Name) Then
                        First_District_ComBobox.Items.Add(Name)
                        Second_District_ComBobox.Items.Add(Name) ' adding the district into the combo boxes 
                    End If
                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If
    End Sub

    Private Sub District_Management_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' this sub procedure calls the procedure that displays all the districts details 
        Display_Districts_Table()
        Delete_District_Button.Enabled = False
        Update_District_Button.Enabled = False
    End Sub

    Private Sub Load_Areas()
        ' this sub procedure loads all the areas into a combo box so it can be selected 
        If DatabaseConnection() Then

            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select AREA_NAME From AREA_TABLE"

                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim Area As String ' this will hold the name of the area read by the recordset 
                    Area = DataReader("AREA_NAME")
                    Add_Area_ComboBox.Items.Add(Area) ' adds to combo box
                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If

    End Sub

    Private Sub Load_District_Name_Click(sender As Object, e As EventArgs) Handles Load_District_Name_Button.Click
        ' this sub procedure occurs when the user clicks the load button. it calls the function "Load_District_Name"
        If District_ID_Listbox.SelectedItem IsNot Nothing Then
            Load_District_Name()
        End If
    End Sub
    Private Sub Load_District_Name()
        'this sub procedure selects the name and the ID of the district pastor that corrosponds to the ID that was selected by the user  
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select DISTRICT_NAME,DISTRICT_PASTOR_ID From DISTRICT_TABLE " & "Where DISTRICT_ID = @DistrictID"
                .Parameters.AddWithValue("@District", Selected_District_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader() 'This is used to run a select query which may return one or more records. It places those records into a recordset variable
                While DataReader.Read
                    Dim Name As String 'holds the name of the district the recordset will read from
                    Name = DataReader("DISTRICT_NAME")
                    UnAssignedPastorID = Int(DataReader("DISTRICT_PASTOR_ID"))
                    Display_District_Name_Textbox.Text = Name ' showing the name in the textbox
                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If
        Delete_District_Button.Enabled = True
    End Sub
    Private Sub Delete_District()
        'this is the actual sub procedure that deletes the district from the database, where the ID is the same as the one selected
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Delete * " & "From DISTRICT_TABLE " & "Where DISTRICT_ID = @DeleteDistrictID " ' the sql command to delete the district
                .Parameters.AddWithValue("@DeleteDistrictID", Selected_District_ID)

                .ExecuteNonQuery() 'This is used to run a query which doesn’t return any result.
            End With
            Connection.Close()
            Display_Districts_Table()
            Display_District_Name_Textbox.Clear() ' clears the text box that had the name of the district 
            Delete_District_Button.Enabled = False
        End If
    End Sub
    Private Sub Assign_New_District_ID()
        ' this sub procedure selects a new district id that will replace the district that is about to be deleted 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select DISTRICT_ID,number From DISTRICT_TABLE Where DISTRICT_ID <> @OldDistrictID" ' where the district id is not the same as the district id that is about to be deleted
                .Parameters.AddWithValue("@OldDistrict", Selected_District_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    New_District_ID = DataReader("DISTRICT_ID") ' storing the new id for the district
                    TemporalDistrictNumber = Int(DataReader("number")) ' storing the number of locals within the district 
                End While
                DataReader.Close()
            End With
        End If
    End Sub
    Private Sub Delete_District_Button_Click(sender As Object, e As EventArgs) Handles Delete_District_Button.Click
        ' this sub procedure is used to verify if the user is adamant on deleting the district
        Dim Delete As String ''creating the local variable delete
        Delete = MsgBox("Are you sure you would like to delete this District? The District's Data will be permenantly deleted, the locals that were assigned to this district will be assigned to another district temporarily(untill you change it)", vbExclamation + vbYesNo + vbDefaultButton2, "Delete District Confirmation")

        If Delete = vbYes Then ' if yes has been selected 
            Assign_New_District_ID()
            GetDeletedDistrictNumber()
            If DatabaseConnection() Then
                Dim SQLCMD As New OleDbCommand
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Update LOCAL_TABLE " & " SET DISTRICT_ID = @NewDistrictID " & "Where DISTRICT_ID = @OldDistrictID" ' update the locals that had this district id with the new district id
                    .Parameters.AddWithValue("@NewDistrict", New_District_ID)
                    .Parameters.AddWithValue("@OldDistrict", Selected_District_ID)
                    .ExecuteNonQuery()
                    UnAssign_Leader()
                    GetNumberToDeleteFromArea()
                    AddToTemporalDistrict()
                End With

                Delete_District()
            End If
        ElseIf Delete = vbNo Then ' if no is selected 
            Display_District_Name_Textbox.Clear() ' clear the textbox
            Delete_District_Button.Enabled = False
        End If
    End Sub
    Private Sub AddToTemporalDistrict()
        'this sub procedure adds the former number of locals in the district that is being deleted and adds it to the temporal district 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Update DISTRICT_TABLE SET [number] = @NewNumber " & "Where DISTRICT_ID = @NewDistrictID"
                .Parameters.AddWithValue("@NewNumber", Int(TemporalDistrictNumber + RemoveNumberInDistricts)) ' adding the old district locals to the new district locals
                .Parameters.AddWithValue("@NewLocalID", New_District_ID)
                .ExecuteNonQuery()
            End With
        End If
    End Sub

    Private Sub GetNumberToDeleteFromArea()
        'this sub procedure obtains the number of districts within the area that the district belonged to and also the area id
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select AREA_TABLE.number,AREA_TABLE.AREA_ID,DISTRICT_ID From AREA_TABLE,DISTRICT_TABLE " & "Where DISTRICT_TABLE.AREA_ID = AREA_TABLE.AREA_ID and DISTRICT_TABLE.DISTRICT_ID = @ID"
                .Parameters.AddWithValue("@ID", Selected_District_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    SubtractNumberInAreas = Int(DataReader("number"))
                    SubtractNumberInAreasID = Int(DataReader("AREA_ID")) ' storing the number and the area id 
                End While
                DataReader.Close()
            End With
            SubtractFromArea() ' calling the procedure that deducts 1 district from te area

        End If
    End Sub

    Private Sub GetDeletedDistrictNumber()
        'this sub procedure stores the number of locals the deleted district had
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select DISTRICT_ID,number From DISTRICT_TABLE Where DISTRICT_ID = @OldDistrictID"
                .Parameters.AddWithValue("@OldDistrictID", Selected_District_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    RemoveNumberInDistricts = Int(DataReader("number")) ' storing the number
                End While
                DataReader.Close()
            End With
        End If
    End Sub
    Private Sub Load_Assigned_District_Pastors()
        'this sub procedure loads the names of district pastors into a combo box so the user can select a pastor 
        If DatabaseConnection() Then

            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select MEMBER_NAME,MEMBER_ID,DISTRICT_PASTOR_ID From MEMBERS_TABLE,DISTRICT_TABLE Where MEMBER_ROLE = @DistrictPastor and CURRENTLY_LEADING = @Leading "
                .Parameters.AddWithValue("@DistrictPastor", "District Pastor")
                .Parameters.AddWithValue("@Leading", "No")

                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim DistrictPastor As String ' this will store the name read by the record set
                    DistrictPastor = DataReader("MEMBER_NAME")

                    If Not New_District_Pastor_ComboBox.Items.Contains(DistrictPastor) Then
                        New_District_Pastor_ComboBox.Items.Add(DistrictPastor) ' addding it into the combobox 
                    End If

                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If
    End Sub
    Private Sub Load_New_Areas()
        'this sub procedure loads the the names of the area into a combo box so the user can select an area 
        If DatabaseConnection() Then

            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select AREA_NAME From AREA_TABLE"
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim Area As String ' stores the name read by the record set 
                    Area = DataReader("AREA_NAME")
                    New_District_Area_ComboBox.Items.Add(Area) ' adding the area into the combo box
                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If
    End Sub

    Private Sub Load_Districts_Details_Button_Click(sender As Object, e As EventArgs) Handles Load_Districts_Details_Button.Click
        ' when this button is clicked it calls these sub procedures 
        If District_ID_Listbox.SelectedItem IsNot Nothing Then
            Load_New_Areas()
            Load_District_Details()
            Load_Assigned_District_Pastors()
        End If
    End Sub
    Private Sub Load_District_Details()
        'this sub procedure loads the details of the selected district into the text boxes so its information can be modifyied where the district's ID is the same as the one that was selected by the user
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select DISTRICT_NAME,DISTRICT_LOCATION,AREA_ID,DISTRICT_PASTOR_ID " & "From DISTRICT_TABLE " & "Where DISTRICT_ID = @DistrictID " 'and LOCAL_TABLE.LOCAL_ID = MEMBERS_TABLE.LOCAL_ID "

                .Parameters.AddWithValue("@DistrictID", Selected_District_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim Name, Location As String ' this will store the values that will be read by the record set 
                    Name = DataReader("DISTRICT_NAME")
                    Location = DataReader("DISTRICT_LOCATION")
                    Area_ID = DataReader("AREA_ID")
                    UnAssignedPastorID = Int(DataReader("DISTRICT_PASTOR_ID"))
                    New_District_Name_TextBox.Text = Name
                    New_District_Location_TextBox.Text = Location ' displaying the info

                End While
                DataReader.Close()
            End With
            UnAssign_Leader()
            Connection.Close() ' closing the connection
            Update_District_Button.Enabled = True
        End If
    End Sub

    Private Sub Update_District_Button_Click(sender As Object, e As EventArgs) Handles Update_District_Button.Click
        ' this subprocedure is called when the update button is clicked. this carries out validation checks on the data the user entered
        If New_District_Name_TextBox.Text = String.Empty Then ' Validation checks on the districts name 
            MsgBox("Please fill in the district's name")
        ElseIf IsNumeric(New_District_Name_TextBox.Text) Then
            MsgBox("numbers in names are not permitted")
        ElseIf Len(New_District_Name_TextBox.Text) < 4 Or Len(New_District_Name_TextBox.Text) > 15 Then
            MsgBox("Please enter a name that is between 4 and 15 characters")


        ElseIf New_District_Location_TextBox.Text = String.Empty Then 'Validation checks on the districts location
            MsgBox("Please fill in the districts location")
        ElseIf IsNumeric(New_District_Location_TextBox.Text) Then
            MsgBox("numbers are not permitted in a location")

        ElseIf Len(New_District_Location_TextBox.Text) < 4 Or Len(New_District_Location_TextBox.Text) > 15 Then
            MsgBox("Please enter a location that is between 4 and 15 characters")


        ElseIf New_District_Pastor_ComboBox.SelectedItem Is Nothing Then 'Validation to check if the district pastor has been selected 
            MsgBox("Please select a district pastor")
        ElseIf New_District_Area_ComboBox.SelectedItem Is Nothing Then 'Validation to check if an area has been chosen 
            MsgBox("Please select an area")
        Else
            Update_District_Details()
        End If
    End Sub

    Private Sub Update_District_Details()
        'this sub procedure updates the district that corresponds to the district ID that was selected by the user
        If DatabaseConnection() Then
            UpdateNumbersInArea() ' calls the sub procedure bellow 
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Update DISTRICT_TABLE " & "Set DISTRICT_NAME = @DistrictName, " & "DISTRICT_LOCATION = @Location, " & "DISTRICT_PASTOR_ID = @DistrictPastor, " & "AREA_ID = @AreaID,  " & "YEARS_ACTIVE = @years " & "Where DISTRICT_ID = @DistrictID "

                .Parameters.AddWithValue("@DistrictName", New_District_Name_TextBox.Text)
                .Parameters.AddWithValue("@Location", New_District_Location_TextBox.Text)
                .Parameters.AddWithValue("@DistrictPastor", New_District_Pastor_ID)
                .Parameters.AddWithValue("@AreaID", New_Area_ID)
                .Parameters.AddWithValue("@years", NewYearsActive)
                .Parameters.AddWithValue("@DistrictID", Selected_District_ID)
                .ExecuteNonQuery()
                'updates the districts details 
                Update_District_Button.Enabled = False

            End With

            Connection.Close()
            Assign_New_Leader()
            Display_Districts_Table() 'show the details of the districts

        End If
    End Sub
    Private Sub UpdateNumbersInArea()
        ' this sub procedure is used to determine how the area will be updated, whether they will lose a district or not when a new district has been updated
        If DatabaseConnection() Then
            If Area_ID <> New_Area_ID Then ' if the old area id for the district is not the same as the new area id

                Dim SQLCMD As New OleDbCommand
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Select NUMBER,AREA_ID From AREA_TABLE " & "Where AREA_TABLE.AREA_ID = @AreaID"
                    .Parameters.AddWithValue("@AreaID", Area_ID)
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read

                        SubtractNumberInAreasID = Int(DataReader("AREA_ID"))
                        SubtractNumberInAreas = Int(DataReader("NUMBER"))
                        ' storing the id and the number currently in the area
                    End While
                    DataReader.Close()
                End With
                SubtractFromArea()
                GetNumberFromArea()

            End If
        End If

    End Sub

    Private Sub SubtractFromArea()
        ' this sub procedure reduces the number of districts in an area by 1
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Update AREA_TABLE Set [NUMBER] = @Num " & "Where AREA_ID = @ID"

                .Parameters.AddWithValue("@Num", SubtractNumberInAreas - 1) 'subtracting 1 from the current number of districts within the area 

                .Parameters.AddWithValue("@ID", SubtractNumberInAreasID)

                .ExecuteNonQuery()
            End With
        End If

    End Sub

    Private Sub GetNumberFromArea()
        'this procedure obtains the number of ddistricts in the new area the district has been assigned to
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select number From AREA_TABLE " & "Where AREA_TABLE.AREA_ID = @AreaID"
                .Parameters.AddWithValue("@AreaID", New_Area_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    AddNumberInNewArea = Int(DataReader("number")) ' storing the number 
                End While
                DataReader.Close()
            End With
            AddNumberToNewArea()

        End If
    End Sub

    Private Sub AddNumberToNewArea()
        ' this sub procedure increments the number in the new area the district has been assigned to by 1
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Update AREA_TABLE Set [NUMBER] = @Num " & "Where AREA_ID = @ID"
                .Parameters.AddWithValue("@Num", AddNumberInNewArea + 1) ' adding 1  onto the current number 
                .Parameters.AddWithValue("@ID", New_Area_ID)

                .ExecuteNonQuery()
            End With
        End If



    End Sub

    Private Sub New_District_Area_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles New_District_Area_ComboBox.SelectedIndexChanged
        ' this sub procedure slects the new area id that corrosponds to the area the user selected 
        Dim AreaName As String '  holds the name selected from the combo box by the user 
        If New_District_Area_ComboBox.SelectedItem IsNot Nothing Then
            AreaName = New_District_Area_ComboBox.SelectedItem.ToString
            If DatabaseConnection() Then
                Dim SQLCMD As New OleDbCommand
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Select AREA_ID From AREA_TABLE " & "Where AREA_TABLE.AREA_NAME = @AreaName"
                    .Parameters.AddWithValue("@AreaName", AreaName)
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read
                        New_Area_ID = Int(DataReader("AREA_ID")) ' storing the new area id  for the district 

                    End While
                    DataReader.Close()
                End With
                Connection.Close()
            End If

        End If
    End Sub



    Private Sub Search_District_Name_Click(sender As Object, e As EventArgs) Handles Search_District_Name_Button.Click
        'this sub procedure is called when the search by name button is clicked. it carries out validation checks on the name entered by the use
        Dim District As String
        If Search_For_District_Name_TextBox.Text = String.Empty Then ' validation checks on the name
            MsgBox("Please fill in the districts name")
        ElseIf IsNumeric(Search_For_District_Name_TextBox.Text) Then
            MsgBox("numbers in names are not permitted")
        Else
            District = Search_For_District_Name_TextBox.Text
            Search_District_Name(District) ' calls the sub procedure 
        End If
        
    End Sub


    Private Sub Search_District_Name(District As String) ' the passed in parameter 
        ' this sub procedure searches the database for a district that matches the name entered or a similar name
        If DatabaseConnection() Then
            Clear_Fields()
            Dim SQLcmd As New OleDbCommand
            With SQLcmd
                .Connection = Connection
                .CommandText = "Select *,MEMBER_NAME,AREA_NAME " & "From DISTRICT_TABLE,MEMBERS_TABLE,AREA_TABLE " & " Where DISTRICT_TABLE.DISTRICT_PASTOR_ID = MEMBERS_TABLE.MEMBER_ID and DISTRICT_NAME Like @NameDistrict and DISTRICT_TABLE.AREA_ID = AREA_TABLE.AREA_ID "

                .Parameters.AddWithValue("@NameDistrict", "%" & District & "%") ' if there is a similar name 
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim DistrictID, Name, Location, DistrictPastor, Area As String ' these variables will store the values read by the record set
                    DistrictID = DataReader("DISTRICT_ID")
                    Name = DataReader("DISTRICT_NAME")
                    Location = DataReader("DISTRICT_LOCATION")
                    Area = DataReader("AREA_NAME")
                    DistrictPastor = DataReader("MEMBER_NAME")

                    District_ID_Listbox.Items.Add(DistrictID)
                    District_Name_Listbox.Items.Add(Name)
                    District_Location_Listbox.Items.Add(Location)
                    District_Pastor_Listbox.Items.Add(DistrictPastor)
                    District_Area_Listbox.Items.Add(Area) ' displaying the records 

                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If
    End Sub

    Private Sub Search_District_ID_Button_Click(sender As Object, e As EventArgs) Handles Search_District_ID_Button.Click
        'this sub procedure is called when the search by ID button is clicked. it carries out validation checks on the ID entered by the user
        Dim ID As Integer
        If Search_For_District_ID_TextBox.Text = String.Empty Then ' validation checks 
            MsgBox("Please fill in the district's ID")
        ElseIf Not IsNumeric(Search_For_District_ID_TextBox.Text) Then
            MsgBox("strings are not permitted in an ID")
        Else
            ID = Int(Search_For_District_ID_TextBox.Text)
            Search_District_ID(ID) 'passes the parameter through
        End If
        
    End Sub

    Private Sub Search_District_ID(District_ID As Integer) ' the passed parameter 
        'this sub procedure searches the database for the record that corresponds to the id that the user typed in the text box
        If DatabaseConnection() Then
            Clear_Fields()
            Dim SQLcmd As New OleDbCommand
            With SQLcmd
                .Connection = Connection
                .CommandText = "Select *,MEMBER_NAME,AREA_NAME " & "From DISTRICT_TABLE,MEMBERS_TABLE,AREA_TABLE " & " Where DISTRICT_TABLE.DISTRICT_PASTOR_ID = MEMBERS_TABLE.MEMBER_ID and DISTRICT_ID = @DistrictID and DISTRICT_TABLE.AREA_ID = AREA_TABLE.AREA_ID " ' and LOCAL_TABLE.LOCAL_ID = MEMBERS_TABLE.LOCAL.ID " 'This is very similar to the StudentID Select query but this time we’re searching using the Surname field and the recordset returned can contain multiple records

                .Parameters.AddWithValue("@DistrictID", District_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim DistrictID, Name, Location, DistrictPastor, Area As String ''stores the info read by the recordset
                    DistrictID = DataReader("DISTRICT_ID")
                    Name = DataReader("DISTRICT_NAME")
                    Location = DataReader("DISTRICT_LOCATION")
                    Area = DataReader("AREA_NAME")
                    DistrictPastor = DataReader("MEMBER_NAME")

                    District_ID_Listbox.Items.Add(DistrictID)
                    District_Name_Listbox.Items.Add(Name)
                    District_Location_Listbox.Items.Add(Location)
                    District_Pastor_Listbox.Items.Add(DistrictPastor)
                    District_Area_Listbox.Items.Add(Area) ' displaying the info

                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If
    End Sub

    Private Sub Add_Area_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Add_Area_ComboBox.SelectedIndexChanged
        ' this  sub procedure stoes the area id an the number of districts within the area that is chosen
        Dim AreaName As String ' this will store the name that was selected from the combo box 
        If Add_Area_ComboBox.SelectedItem IsNot Nothing Then
            AreaName = Add_Area_ComboBox.SelectedItem.ToString
            If DatabaseConnection() Then
                Dim SQLCMD As New OleDbCommand
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Select AREA_ID,NUMBER From AREA_TABLE " & "Where AREA_TABLE.AREA_NAME = @AreaName"
                    .Parameters.AddWithValue("@AreaName", AreaName)
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read
                        Area_ID = Int(DataReader("AREA_ID"))
                        AddNumberInAreas = Int(DataReader("NUMBER")) ' storing the data
                    End While
                    DataReader.Close()
                End With
                Connection.Close()
            End If

        End If
    End Sub


    Private Sub Search_District_Pastor_Button_Click(sender As Object, e As EventArgs) Handles Search_District_Pastor_Button.Click
        'this sub procedure is called when the search by name button is clicked. it carries out validation checks on the name entered by the use
        Dim PastorName As String
        If Search_For_District_Pastor_Name_TextBox.Text = String.Empty Then ' validation checks on the name
            MsgBox("Please fill in the district pastor name")
        ElseIf IsNumeric(Search_For_District_Pastor_Name_TextBox.Text) Then
            MsgBox("numbers in names are not permitted")
        Else
            PastorName = Search_For_District_Pastor_Name_TextBox.Text
            Search_District_Pastor_Name(PastorName) ' calls the sub procedure 
        End If
    End Sub
    Private Sub Search_District_Pastor_Name(DistrictPastorName As String)
        'this sub procedure searches the database for the record that corresponds to the name that the user typed in the text box
        If DatabaseConnection() Then
            Clear_Fields()
            Dim SQLcmd As New OleDbCommand
            With SQLcmd
                .Connection = Connection
                .CommandText = "Select *,MEMBER_NAME,AREA_NAME " & "From DISTRICT_TABLE,MEMBERS_TABLE,AREA_TABLE " & " Where DISTRICT_TABLE.DISTRICT_PASTOR_ID = MEMBERS_TABLE.MEMBER_ID and MEMBER_NAME Like @Name and DISTRICT_TABLE.AREA_ID = AREA_TABLE.AREA_ID " ' and LOCAL_TABLE.LOCAL_ID = MEMBERS_TABLE.LOCAL.ID " 'This is very similar to the StudentID Select query but this time we’re searching using the Surname field and the recordset returned can contain multiple records

                .Parameters.AddWithValue("@Name", "%" & DistrictPastorName & "%") '
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim DistrictID, Name, Location, DistrictPastor, Area As String ' these will hold the data that is read
                    DistrictID = DataReader("DISTRICT_ID")
                    Name = DataReader("DISTRICT_NAME")
                    Location = DataReader("DISTRICT_LOCATION")
                    Area = DataReader("AREA_NAME")
                    DistrictPastor = DataReader("MEMBER_NAME")
                    District_ID_Listbox.Items.Add(DistrictID)
                    District_Name_Listbox.Items.Add(Name)
                    District_Location_Listbox.Items.Add(Location)
                    District_Pastor_Listbox.Items.Add(DistrictPastor)
                    District_Area_Listbox.Items.Add(Area) ' display the data

                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If
    End Sub

    Private Sub First_District_ComBobox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles First_District_ComBobox.SelectedIndexChanged
        ' this sub selects the area id and the amount of years the selected district has 
        Dim DistrictName As String ' stores the district name
        If First_District_ComBobox.SelectedItem IsNot Nothing Then
            DistrictName = First_District_ComBobox.SelectedItem.ToString
            If DatabaseConnection() Then
                Dim SQLCMD As New OleDbCommand
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Select DISTRICT_ID,AREA_ID,YEARS_ACTIVE From DISTRICT_TABLE " & "Where DISTRICT_TABLE.DISTRICT_NAME = @DistrictName"
                    .Parameters.AddWithValue("@DistrictName", DistrictName)
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read
                        SecondDistrictID = Int(DataReader("DISTRICT_ID"))
                        FirstAreaID = Int(DataReader("AREA_ID"))
                        FirstDistrictPastorYears = Int(DataReader("YEARS_ACTIVE")) ' storing the data 
                    End While
                    DataReader.Close()
                End With
                Connection.Close()
                Load_Second_District_Pastor()
            End If

        End If

    End Sub

    Private Sub Second_District_ComBobox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Second_District_ComBobox.SelectedIndexChanged
        ' this sub selects the area id and the amount of years the selected district has 
        Dim DistrictName As String
        If Second_District_ComBobox.SelectedItem IsNot Nothing Then
            DistrictName = Second_District_ComBobox.SelectedItem.ToString
            'MsgBox(AreaName)
            If DatabaseConnection() Then
                Dim SQLCMD As New OleDbCommand
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Select DISTRICT_ID,AREA_ID,YEARS_ACTIVE From DISTRICT_TABLE " & "Where DISTRICT_TABLE.DISTRICT_NAME = @DistrictName"
                    .Parameters.AddWithValue("@DistrictName", DistrictName)
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read
                        FirstDistrictID = Int(DataReader("DISTRICT_ID"))
                        SecondAreaID = Int(DataReader("AREA_ID"))
                        SecondDistrictPastorYears = Int(DataReader("YEARS_ACTIVE")) ' storing the data 
                    End While
                    DataReader.Close()
                End With
                Connection.Close()
                Load_First_District_Pastor()
            End If

        End If

    End Sub
    Private Sub Load_Second_District_Pastor()
        ' this sub displays the details of the district pastor that is about to be transfered
        If DatabaseConnection() Then
            Dim SQLcmd As New OleDbCommand
            With SQLcmd
                .Connection = Connection
                .CommandText = "Select DISTRICT_PASTOR_ID,DISTRICT_ID,MEMBER_ID,MEMBER_NAME,YEARS_ACTIVE " & "From DISTRICT_TABLE,MEMBERS_TABLE " & "Where DISTRICT_TABLE.DISTRICT_PASTOR_ID = MEMBERS_TABLE.MEMBER_ID and DISTRICT_ID = @DistrictID " ' and LOCAL_TABLE.LOCAL_ID = MEMBERS_TABLE.LOCAL.ID " 'This is very similar to the StudentID Select query but this time we’re searching using the Surname field and the recordset returned can contain multiple records

                .Parameters.AddWithValue("@DistrictID", SecondDistrictID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim DistrictPastor, years As String ' these variables will store the data read from the record set
                    SecondDistrictPastorID = Int(DataReader("DISTRICT_PASTOR_ID"))
                    DistrictPastor = DataReader("MEMBER_NAME")
                    years = DataReader("YEARS_ACTIVE")
                    First_District_Pastor_Name_Textbox.Text = DistrictPastor
                    Display_First_Years_Textbox.Text = "years active: " & years ' displaying the data

                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If
    End Sub
    Private Sub Load_First_District_Pastor()
        ' this sub displays the details of the district pastor that is about to be transfered
        If DatabaseConnection() Then
            Dim SQLcmd As New OleDbCommand
            With SQLcmd
                .Connection = Connection
                .CommandText = "Select DISTRICT_PASTOR_ID,DISTRICT_ID,YEARS_ACTIVE,MEMBER_ID,MEMBER_NAME " & "From DISTRICT_TABLE,MEMBERS_TABLE " & "Where DISTRICT_TABLE.DISTRICT_PASTOR_ID = MEMBERS_TABLE.MEMBER_ID and DISTRICT_ID = @DistrictID " ' and LOCAL_TABLE.LOCAL_ID = MEMBERS_TABLE.LOCAL.ID " 'This is very similar to the StudentID Select query but this time we’re searching using the Surname field and the recordset returned can contain multiple records

                .Parameters.AddWithValue("@DistrictID", FirstDistrictID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read

                    Dim DistrictPastorName, years As String ' these variables will store the data read from the record set
                    FirstDistrictPastorID = Int(DataReader("DISTRICT_PASTOR_ID"))
                    DistrictPastorName = DataReader("MEMBER_NAME")
                    years = DataReader("YEARS_ACTIVE")
                    Second_District_Pastor_Name_Textbox.Text = DistrictPastorName
                    Display_Second_Years_Textbox.Text = "years active: " & years

                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If
    End Sub
    Private Sub Transfer_District_Pastor_Button_Click(sender As Object, e As EventArgs) Handles Transfer_District_Pastor_Button.Click
        ' this sub procedure checks if all the conditions are met in order to transfer a district pastor 
        If FirstAreaID = SecondAreaID And (FirstDistrictPastorYears >= 6 And SecondDistrictPastorYears >= 6) Then
            Transfer_First_District_Pastor()
            Transfer_Second_District_Pastor()
            Display_Districts_Table()
        ElseIf FirstDistrictPastorID = SecondDistrictPastorID Then ' same pastor selected 
            MsgBox("Select 2 Different Districts with different District Pastors")
        ElseIf SecondAreaID <> FirstAreaID Then ' based in different areas 
            MsgBox("Pastors that have been selected needs to be within the same Area")
        ElseIf FirstDistrictPastorYears < 6 Or SecondDistrictPastorYears < 6 Then
            MsgBox("Pastors can not be transfered if their time has not exceeded 6 years")
        End If

    End Sub

    Private Sub Transfer_First_District_Pastor()
        ' this sub procedure updates the district and sets the district pastor id to the other district pastor id that belonged to the other district
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Update DISTRICT_TABLE " & "Set DISTRICT_PASTOR_ID = @NewDistrictPastorID " & "Where DISTRICT_ID = @DistrictID"
                .Parameters.AddWithValue("@NewDistrictPastorID", FirstDistrictPastorID) ' assigning the other pastor  id 
                .Parameters.AddWithValue("@DistrictID", SecondDistrictID)
                .ExecuteNonQuery()
            End With
            Connection.Close()

        End If
    End Sub
    Private Sub Transfer_Second_District_Pastor()
        ' this sub procedure updates the district and sets the district pastor id to the other district pastor id that belonged to the other district
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Update DISTRICT_TABLE " & "Set DISTRICT_PASTOR_ID = @NewDistrictPastorID " & "Where DISTRICT_ID = @DistrictID"
                .Parameters.AddWithValue("@NewDistrictPastorID", SecondDistrictPastorID) ' assigning the other pastor  id 
                .Parameters.AddWithValue("@DistrictID", FirstDistrictID)
                .ExecuteNonQuery()
            End With
            Connection.Close()
        End If
    End Sub

    Private Sub Add_District_Pastor_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Add_District_Pastor_ComboBox.SelectedIndexChanged
        ' this sub procedure selects the corrosponding pastor id that matches the name of the pastor selected by the user
        Dim Pastor As String '
        If Add_District_Pastor_ComboBox.SelectedItem IsNot Nothing Then
            Pastor = Add_District_Pastor_ComboBox.SelectedItem.ToString
            If DatabaseConnection() Then
                Dim SQLCMD As New OleDbCommand
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Select MEMBER_ID From MEMBERS_TABLE " & "Where MEMBERS_TABLE.MEMBER_NAME = @DistrictPastorName"
                    .Parameters.AddWithValue("@DistrictPastorName", Pastor)
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read
                        District_Pastor_ID = Int(DataReader("MEMBER_ID")) ' storing the id

                    End While
                    DataReader.Close()
                End With
                Connection.Close()
            End If

        End If
    End Sub


    Private Sub New_District_Pastor_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles New_District_Pastor_ComboBox.SelectedIndexChanged
        ' this sub procedure selects the corrosponding pastor id that matches the name of the new pastor selected by the user
        Dim Pastor As String
        If New_District_Pastor_ComboBox.SelectedItem IsNot Nothing Then
            Pastor = New_District_Pastor_ComboBox.SelectedItem.ToString
            If DatabaseConnection() Then
                Dim SQLCMD As New OleDbCommand
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Select MEMBER_ID From MEMBERS_TABLE " & "Where MEMBERS_TABLE.MEMBER_NAME = @DistrictPastor"
                    .Parameters.AddWithValue("@DistrictPastor", Pastor)
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read
                        New_District_Pastor_ID = Int(DataReader("MEMBER_ID")) ' storing the id

                    End While
                    DataReader.Close()
                End With
                Connection.Close()
            End If

        End If
    End Sub

    Private Sub District_Reports_Button_Click(sender As Object, e As EventArgs) Handles District_Reports_Button.Click
        ' thsi sub calls the ranking district form to show
        District_Ranking_Form.ShowDialog()
    End Sub

    Private Sub District_Years_NumericUpDown_ValueChanged(sender As Object, e As EventArgs) Handles District_Years_NumericUpDown.ValueChanged
        ' this sub stores the years from the numeric updaown
        CurrentYearsActive = Int(District_Years_NumericUpDown.Value.ToString)
    End Sub

    Private Sub Update_District_Years_NumericUpDown_ValueChanged(sender As Object, e As EventArgs) Handles Update_District_Years_NumericUpDown.ValueChanged
        ' this sub stores the years from the numeric updaown
        NewYearsActive = Int(Update_District_Years_NumericUpDown.Value.ToString)

    End Sub

    Private Sub Add_District_Clear_Button_Click(sender As Object, e As EventArgs) Handles Add_District_Clear_Button.Click
        'Clears inputs
        District_Name_TextBox.Clear()
        District_Location_TextBox.Clear()
        Add_Area_ComboBox.SelectedIndex = -1
        Add_District_Pastor_ComboBox.SelectedIndex = -1
        District_Years_NumericUpDown.Value = 0
    End Sub

    Private Sub Search_District_Clear_Button_Click(sender As Object, e As EventArgs) Handles Search_District_Clear_Button.Click
        'Clears inputs
        Search_For_District_ID_TextBox.Clear()
        Search_For_District_Name_TextBox.Clear()
    End Sub

    Private Sub Search_District_Pastor_Clear_Button_Click(sender As Object, e As EventArgs) Handles Search_District_Pastor_Clear_Button.Click
        'Clears inputs
        Search_For_District_Pastor_Name_TextBox.Clear()
    End Sub

    Private Sub Update_District_Clear_Button_Click(sender As Object, e As EventArgs) Handles Update_District_Clear_Button.Click
        'Clears inputs
        New_District_Name_TextBox.Clear()
        New_District_Location_TextBox.Clear()
        New_District_Area_ComboBox.SelectedIndex = -1
        New_District_Pastor_ComboBox.SelectedIndex = -1
        Update_District_Years_NumericUpDown.Value = 0
    End Sub
End Class