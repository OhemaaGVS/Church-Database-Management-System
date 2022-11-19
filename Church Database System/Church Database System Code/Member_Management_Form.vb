Imports System.Data.OleDb
Public Class Member_Management_Form
    Private AddNumberInLocals As Integer ' this variable is used to store the amount of members who attend one local
    Private AddNumberInNewLocals As Integer ' this varaible is used to store the current amount of members based at a new local that a member will have been assined to 
    Private SubtractNumberInLocalsID As Integer ' this variable stores the ID of the previos local the member was assined to, so the system knows which local to remove a member from
    Private SubtractNumberInLocals As Integer ' this variable stores the current number of members in a local before one member will be removed
    Private IsALeader As Boolean ' this variable is used to identify if the member is a leader or not( so if the member is an Area head, district pastor or presiding elder)
    Private CurrentlyLeading As String ' this is a variable that holds the value that the system reads from the "CurrentlyLeading" field in the MEMBERS_TABLE in the database
    Private ValueOfMemberID As Integer = -1 ' this  variable holds the generated ID for the member that will be newly created
    Private Local_ID As Integer ' this variable stores the local id that will be assined to the member. this will enable the system to identify what local the member attends
    Private Role As String ' this variable stores the role that will be assined to the member from the combobox
    Private Selected_Member_ID As Integer ' this variable stores the ID that has been selected from the list box that contains the member IDs
    Private NewRole As String ' this variable holds the new role that the member will be assined
    Private NewLocal_ID As Integer ' this variable holds the new Local that the member will be atteneding

    Private Sub Member_ID_Listbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Member_ID_Listbox.SelectedIndexChanged
        '* this sub procedure is used to store the value of the selected member ID from the listbox*
        If Member_ID_Listbox.SelectedItem IsNot Nothing Then
            Selected_Member_ID = Int(Member_ID_Listbox.SelectedItem)
        End If
    End Sub

    Private Sub Add_Member_Button_Click(sender As Object, e As EventArgs) Handles Add_Member_Button.Click
        '* when the button is clicked it will carry out some validation checks to make sure invalid data is not being entered. if the fields have been correctly filled it it calls the sub procedure "Create_New_Member" *
        If Member_Name_TextBox.Text = String.Empty Then ' Validation checks on the members name 
            MsgBox("Please fill in the member's name")
        ElseIf IsNumeric(Member_Name_TextBox.Text) Then
            MsgBox("numbers in names are not permitted")
        ElseIf Len(Member_Name_TextBox.Text) < 4 Or Len(Member_Name_TextBox.Text) > 15 Then
            MsgBox("Please enter a name that is between 4 and 15 characters")

        ElseIf Member_Age_TextBox.Text = String.Empty Then 'Validation checks on the members age 
            MsgBox("Please fill in the member's age")
        ElseIf Not IsNumeric(Member_Age_TextBox.Text) Then
            MsgBox("strings are not permitted in an age")


        ElseIf Member_Telephone_TextBox.Text = String.Empty Then 'Validation checks on the members phone number
            MsgBox("Please fill in the member's phone number")
        ElseIf Not IsNumeric(Member_Telephone_TextBox.Text) Then
            MsgBox("letters are not permitted in a telephone number")

        ElseIf Len(Member_Telephone_TextBox.Text) <> 11 Or Int(Member_Telephone_TextBox.Text.Substring(0, 1)) <> 0 Then
            MsgBox("Please enter a number that is 11 didgits with 0 as the starting number")


        ElseIf Add_Member_Role_ComboBox.SelectedItem Is Nothing Then 'Validation to check if the role of the member has been selected 
            MsgBox("Please select a role for the member")
        ElseIf Add_Member_Local_ComboBox.SelectedItem Is Nothing And IsALeader = False Then 'Validation to check if the local of the member has been selected 
            MsgBox("Please select a local for the member")
        Else
            Create_New_Member()

        End If
    End Sub

    Private Sub Create_New_Member()
        ' * this sub procedure stores the new member that has been created in the database *
        If DatabaseConnection() Then ' checking if there is a database connection 
            Dim SQLCMD As New OleDbCommand 'Open a connection to the database

            If ValueOfMemberID = -1 And IsALeader = False Then ' if the value of the members id is currently -1 and the boolean isaleader is false then this section of code will call the subprocedure "AddToLocal" and will store the members details in the database 


                AddToLocal()
                With SQLCMD
                    .Connection = Connection

                    .CommandText = "Insert into MEMBERS_TABLE (MEMBER_NAME, MEMBER_AGE, MEMBER_TELEPHONE_NUMBER, MEMBER_ROLE, LOCALS_ID, CURRENTLY_LEADING )" & "Values (@MemberName ,@MemberAge ,@MemberTelNum ,@MemberRole ,@LocalID, @Leading )"
                    .Parameters.AddWithValue("@MemberName", Member_Name_TextBox.Text)
                    .Parameters.AddWithValue("@MemberAge", Member_Age_TextBox.Text)
                    .Parameters.AddWithValue("@MemberTelNum", Member_Telephone_TextBox.Text)
                    .Parameters.AddWithValue("@MemberRole", Role)
                    .Parameters.AddWithValue("@LocalID", Local_ID)
                    .Parameters.AddWithValue("@Leading", "No")
                    .ExecuteNonQuery()
                    .CommandText = "Select @@Identity" '
                    ValueOfMemberID = .ExecuteScalar
                    Add_Member_Auto_ID_Label.Text = ValueOfMemberID ' storing the ID assigned to the member

                End With
                Connection.Close()
                Add_Member_Auto_ID_Label.Text = "Automatically Generated"

                Display_Members_Table() ' calling the procedure that displays the information of the members 
            ElseIf ValueOfMemberID = -1 And IsALeader = True Then ' if the mmber id is -1 and the boolean is leader is true then it adds the new member into the database without assigning them a local that they will be based at
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Insert into MEMBERS_TABLE (MEMBER_NAME, MEMBER_AGE, MEMBER_TELEPHONE_NUMBER, MEMBER_ROLE, CURRENTLY_LEADING )" & "Values (@MemberName ,@MemberAge ,@MemberTelNum ,@MemberRole ,@Leading )"
                    .Parameters.AddWithValue("@MemberName", Member_Name_TextBox.Text)
                    .Parameters.AddWithValue("@MemberAge", Member_Age_TextBox.Text)
                    .Parameters.AddWithValue("@MemberTelNum", Member_Telephone_TextBox.Text)
                    .Parameters.AddWithValue("@MemberRole", Role)
                    .Parameters.AddWithValue("@Leading", "No")

                    .ExecuteNonQuery() 'This is used to run a query which doesn’t return any result. 
                    .CommandText = "Select @@Identity"
                    ValueOfMemberID = .ExecuteScalar 'This is used to run a Select query which returns a single value (usually a number). 
                    Add_Member_Auto_ID_Label.Text = ValueOfMemberID

                End With
                Connection.Close()
                Add_Member_Auto_ID_Label.Text = "Automatically Generated"

                Display_Members_Table() ' calling the procedure that displays the information of the members 
            End If

        End If

    End Sub

    Private Sub AddToLocal()
        ' this sub procedure will increase the number of members a local where a new member has been assigned to 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Update LOCAL_TABLE Set [NUMBER] = @Num " & "Where LOCAL_ID = @ID"

                .Parameters.AddWithValue("@Num", AddNumberInLocals + 1) ' adding 1 to the current number the local has

                .Parameters.AddWithValue("@ID", Local_ID)

                .ExecuteNonQuery() 'This is used to run a query which doesn’t return any result. 
            End With
        End If



    End Sub

    Private Sub GetNumberFromLocal()
        'this subprocedure retrives the number of members within the new local that a member has been assined to
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select NUMBER From LOCAL_TABLE " & "Where LOCAL_TABLE.LOCAL_ID = @LocalID"
                .Parameters.AddWithValue("@LocalID", NewLocal_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader() 'This is used to run a select query which may return one or more records. It places those records into a recordset variable
                While DataReader.Read

                    AddNumberInNewLocals = Int(DataReader("NUMBER")) ' storing the number so it can be used later

                End While
                DataReader.Close()
            End With
            AddNumberToNewLocal()

        End If
    End Sub

    Private Sub AddNumberToNewLocal()
        ' this subprocedure incrememnts the number in the new local that the member has been allocated to by 1
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Update LOCAL_TABLE Set [NUMBER] = @Num " & "Where LOCAL_ID = @ID"

                .Parameters.AddWithValue("@Num", AddNumberInNewLocals + 1) ' adding 1 onto the number of members that the local has 
                ' MsgBox(Local_ID)
                .Parameters.AddWithValue("@ID", NewLocal_ID)

                .ExecuteNonQuery() 'This is used to run a query which doesn’t return any result. 
            End With
        End If



    End Sub
    Private Sub SubtractFromLocal()
        ' this subprocedure is used tho decrease the number of members in a local by one when a member is moved to a different local
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Update LOCAL_TABLE Set [NUMBER] = @Num " & "Where LOCAL_ID = @ID"

                .Parameters.AddWithValue("@Num", SubtractNumberInLocals - 1) ' subtracting one from the local that the member used to be part of
                ' MsgBox(Local_ID)
                .Parameters.AddWithValue("@ID", SubtractNumberInLocalsID)

                .ExecuteNonQuery() 'This is used to run a query which doesn’t return any result. 
            End With
        End If

    End Sub
    Private Sub Clear_Fields()
        ' this subprocedures function is tho clear the textboxes, listboxes, comboboxes and set the member id to -1 so that the latest contents of the database can be displayed 
        Member_ID_Listbox.Items.Clear()
        Member_Name_Listbox.Items.Clear()
        Member_Age_Listbox.Items.Clear()
        Member_Number_Listbox.Items.Clear()
        Member_Role_Listbox.Items.Clear()
        Member_Local_Listbox.Items.Clear()
        Member_Name_TextBox.Clear()
        Member_Age_TextBox.Clear()
        Member_Telephone_TextBox.Clear()
        New_Member_Name_TextBox.Clear()
        New_Member_Age_TextBox.Clear()
        Search_Member_Name_TextBox.Clear()
        Search_Member_ID_TextBox.Clear()
        New_Member_Telephone_TextBox.Clear()
        Add_Member_Local_ComboBox.SelectedIndex = -1
        Add_Member_Role_ComboBox.SelectedIndex = -1
        New_Member_Role_ComboBox.SelectedIndex = -1
        New_Member_Local_ComboBox.SelectedIndex = -1
        New_Member_Role_ComboBox.Items.Clear()
        New_Member_Local_ComboBox.Items.Clear()
        Add_Member_Local_ComboBox.Items.Clear()
        Add_Member_Role_ComboBox.Items.Clear()
        ValueOfMemberID = -1
        Display_Member_Name_Textbox.Clear()

    End Sub

    Private Sub Display_Members_Table()
        ' this sub procedure displays the information of the members and the locals they have been assigned to
        If DatabaseConnection() Then ' checking for the connection to the database
            Clear_Fields() ' calling the subprocedure "Clear_Fields"
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select LOCAL_NAME,* From LOCAL_TABLE,MEMBERS_TABLE " & "Where LOCAL_TABLE.LOCAL_ID = MEMBERS_TABLE.LOCALS_ID and  (MEMBERS_TABLE.MEMBER_ROLE <> @Role1 and MEMBERS_TABLE.MEMBER_ROLE <> @Role2  and MEMBERS_TABLE.MEMBER_ROLE <> @Role3)"
                .Parameters.AddWithValue("@Role1", "District Pastor")
                .Parameters.AddWithValue("@Role2", "Presiding Elder")
                .Parameters.AddWithValue("@Role3", "Area Head")
                Dim DataReader As OleDbDataReader = .ExecuteReader() ' 'This is used to run a select query which may return one or more records. It places those records into a recordset variable
                While DataReader.Read ' Data.Read gets the next record from the record set and returns True. If there are no records in the recordset then it returns False.

                    Dim MemberID, Name, Number, Role, Age, Local As String ' creating variables that will store the data that has been read from the data reader 
                    MemberID = DataReader("MEMBER_ID")
                    Name = DataReader("MEMBER_NAME")
                    Age = DataReader("MEMBER_AGE")
                    Number = DataReader("MEMBER_TELEPHONE_NUMBER")
                    Role = DataReader("MEMBER_ROLE")
                    Local = DataReader("LOCAL_NAME")

                    Member_ID_Listbox.Items.Add(MemberID) ' displaying the data that has been read
                    Member_Name_Listbox.Items.Add(Name)
                    Member_Age_Listbox.Items.Add(Age)
                    Member_Number_Listbox.Items.Add(Number)
                    Member_Role_Listbox.Items.Add(Role)
                    Member_Local_Listbox.Items.Add(Local)
                End While
                DataReader.Close() ' closing the recordset
            End With
            Connection.Close() ' closing the connection 
        End If
        Load_Member_Roles() ' calling sub procedures 
        Load_Locals()
        Display_Leaders()
    End Sub
    Private Sub Display_Leaders()
        ' this subprocedure displays the information of the members who have not been assined a local to attend (So area heads, presiding elders and district pastors)
        If DatabaseConnection() Then

            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select * From MEMBERS_TABLE Where MEMBER_ROLE = @Role1 or MEMBER_ROLE = @Role2 or MEMBER_ROLE = @Role3 " ' where their role is either an area head, district pastor or a presiding elder
                .Parameters.AddWithValue("@Role1", "District Pastor")
                .Parameters.AddWithValue("@Role2", "Presiding Elder")
                .Parameters.AddWithValue("@Role3", "Area Head")

                Dim DataReader As OleDbDataReader = .ExecuteReader() 'This is used to run a select query which may return one or more records. It places those records into a recordset variable
                While DataReader.Read
                    Dim MemberID, Name, Number, Role, Age As String ' the variables that will hold the data that will be read by the record set
                    MemberID = DataReader("MEMBER_ID")
                    Name = DataReader("MEMBER_NAME")
                    Age = DataReader("MEMBER_AGE")
                    Number = DataReader("MEMBER_TELEPHONE_NUMBER")
                    Role = DataReader("MEMBER_ROLE")


                    Member_ID_Listbox.Items.Add(MemberID) ' dislplaying the data read into each listbox 
                    Member_Name_Listbox.Items.Add(Name)
                    Member_Age_Listbox.Items.Add(Age)
                    Member_Number_Listbox.Items.Add(Number)
                    Member_Role_Listbox.Items.Add(Role)
                End While
                DataReader.Close() ' closing the record set
            End With
            Connection.Close() ' closing the connection 
        End If

    End Sub

    Private Sub Member_Management_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' this sub procedure is called when the form loads. it calls the "Display Members" sub procedure and disables buttons from other sections of the form 
        Display_Members_Table()
        Delete_Members_Button.Enabled = False
        Update_Member_Button.Enabled = False

    End Sub
    Private Sub Load_Member_Roles()
        ' this sub procedure is used to load the different roles within the church into a combo bocx so that the user can select a role from it
        Add_Member_Role_ComboBox.Items.Add("Sister")
        Add_Member_Role_ComboBox.Items.Add("Brother")
        Add_Member_Role_ComboBox.Items.Add("Deacon")
        Add_Member_Role_ComboBox.Items.Add("Deaconess")
        Add_Member_Role_ComboBox.Items.Add("Presiding Elder")
        Add_Member_Role_ComboBox.Items.Add("Elder")
        Add_Member_Role_ComboBox.Items.Add("Instrumentalist")
        Add_Member_Role_ComboBox.Items.Add("Singer")
        Add_Member_Role_ComboBox.Items.Add("District Pastor")
        Add_Member_Role_ComboBox.Items.Add("Area Head")
        Add_Member_Role_ComboBox.Items.Add("Pastor")
        Add_Member_Role_ComboBox.Items.Add("Area Head's Child")
        Add_Member_Role_ComboBox.Items.Add("District Pastor's Child") ' these are the different roles
        Add_Member_Role_ComboBox.Items.Add("District Pastor's Wife")
        Add_Member_Role_ComboBox.Items.Add("Pastor's Wife")
        Add_Member_Role_ComboBox.Items.Add("Pastor's Child")
        Add_Member_Role_ComboBox.Items.Add("Area Head's Wife")
        Add_Member_Role_ComboBox.Items.Add("Apostle's Wife")
        Add_Member_Role_ComboBox.Items.Add("Apostle's Child ")
        Add_Member_Role_ComboBox.Items.Add("Apostle")
        Add_Member_Role_ComboBox.Items.Add("Youth member")
        Add_Member_Role_ComboBox.Items.Add("Youth Leader")
        Add_Member_Role_ComboBox.Items.Add("Child")
        Add_Member_Role_ComboBox.Items.Add("Sunday School Teacher")


    End Sub

    Private Sub Add_Member_Role_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Add_Member_Role_ComboBox.SelectedIndexChanged
        ' this sub procedure is called when an item is selcted from the combo box
        If Add_Member_Role_ComboBox.SelectedItem IsNot Nothing Then ' checking if an item has been selected or not 
            Role = Add_Member_Role_ComboBox.SelectedItem.ToString ' assigning the role selected from the combo box to the variable "Role"
            If Role = "Area Head" Or Role = "Presiding Elder" Or Role = "District Pastor" Then ' checks if role is equal area head or district pastor or a presiding elder 
                IsALeader = True ' sets the boolean IsALeader to true 
                Add_Member_Local_ComboBox.Visible = False ' makes the combo box that displays locals invisible
            ElseIf Role <> "Area Head" And Role <> "Presiding Elder" And Role <> "District Pastor" Then ' if its not equal to any of them
                IsALeader = False ' sets the boolean IsALeader to false 
                Add_Member_Local_ComboBox.Visible = True ' sets the Boolean to false 
            End If

        End If
    End Sub

    Private Sub Add_Member_Local_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Add_Member_Local_ComboBox.SelectedIndexChanged
        'the subprocedure selects the local id and the number of members that are within the local that is selected from the combobox
        Dim LocalName As String ' creating a local variable called  LocalName. this will hold the name of the local that the user has selected from the combobox
        If Add_Member_Local_ComboBox.SelectedItem IsNot Nothing Then
            LocalName = Add_Member_Local_ComboBox.SelectedItem.ToString ' if there has been a local name selected from the combobox, assign it to LocalName
            If DatabaseConnection() Then ' if there is a connection 
                Dim SQLCMD As New OleDbCommand
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Select LOCAL_ID,NUMBER From LOCAL_TABLE " & "Where LOCAL_TABLE.LOCAL_NAME = @LocalName" ' selecting the Local id and the number of members that corrosponds the local that was selected 
                    .Parameters.AddWithValue("@LocalName", LocalName)
                    Dim DataReader As OleDbDataReader = .ExecuteReader() 'This is used to run a select query which may return one or more records. It places those records into a recordset variable
                    While DataReader.Read
                        Local_ID = Int(DataReader("LOCAL_ID"))
                        AddNumberInLocals = Int(DataReader("NUMBER")) ' storing the data that the recorset read
                    End While
                    DataReader.Close() ' closing the recordset
                End With
                Connection.Close() ' closing the connection
            End If

        End If
    End Sub
    Private Sub Load_Locals()
        ' this sub procedure is used to load the local names into the combo box so it can be selected by the user
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select LOCAL_NAME From LOCAL_TABLE"

                Dim DataReader As OleDbDataReader = .ExecuteReader() 'This is used to run a select query which may return one or more records. It places those records into a recordset variable
                While DataReader.Read
                    Dim LocalName As String ' creating a local varaible to hold the data that the recordset will read 
                    LocalName = DataReader("LOCAL_NAME")
                    Add_Member_Local_ComboBox.Items.Add(LocalName) ' adding the local name to the combo box
                End While
                DataReader.Close()
            End With
            Connection.Close() ' closing the connection 

        End If
    End Sub

    Private Sub Load_Members_Name_Button_Click(sender As Object, e As EventArgs) Handles Load_Members_Name_Button.Click
        ' this sub procedure occurs when the user clicks the load button. it calls the function "Load_Member_Name"
        If Member_ID_Listbox.SelectedItem IsNot Nothing Then
            Load_Member_Name()
        End If
    End Sub
    Private Sub Load_Member_Name()
        'this sub procedure selcects the name and the role of the member that corrosponds to the ID that was selcted by the user  
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select MEMBER_NAME,MEMBER_ROLE From MEMBERS_TABLE " & "Where MEMBER_ID = @MemberID"
                .Parameters.AddWithValue("@MemberID", Selected_Member_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader() 'This is used to run a select query which may return one or more records. It places those records into a recordset variable
                While DataReader.Read
                    Dim MemberName As String ' local variable Member name holds the name of the member the recordset will read from
                    MemberName = DataReader("MEMBER_NAME") ' assining the data to the local variable
                    Role = DataReader("MEMBER_ROLE") ' storing their role 
                    Display_Member_Name_Textbox.Text = MemberName ' showing the name in the textbox
                End While
                DataReader.Close() ' close record set 
            End With
            Connection.Close()
        End If
        Delete_Members_Button.Enabled = True ' enabling the delete button 
    End Sub
    Private Sub Delete_Member()
        'this is the actual sub procedure that deletes the member from the database, where the ID is the same as the one selected 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Delete * " & "From MEMBERS_TABLE " & "Where MEMBER_ID = @DeleteMemberID " ' the sql command to delete the member 
                .Parameters.AddWithValue("@DeleteMemberID", Selected_Member_ID)
                .ExecuteNonQuery() 'This is used to run a query which doesn’t return any result. 
            End With
            Connection.Close()
            Display_Members_Table()
            Display_Member_Name_Textbox.Clear() ' clears the text box that had the name of the member
            Delete_Members_Button.Enabled = False
        End If
    End Sub

    Private Sub Delete_Members_Button_Click(sender As Object, e As EventArgs) Handles Delete_Members_Button.Click
        'this sub procedure is called when the delete button is clicked. when the button is clicked it produces an error message, if yes is pressed thn it calls the sub procedure to delete the member 
        Dim Delete As String ' creating the local variable delete
        Delete = MsgBox("Are you sure you would like to delete this Member? Data will be permenantly deleted", vbExclamation + vbYesNo + vbDefaultButton2, "Delete Member Confirmation") 'message box alerts the user this data will be permenatly lost
        If Delete = vbYes Then ' if yess is chosen 
            CheckForLeader()
            If (Role <> "Area Head" And Role <> "District Pastor" And Role <> "Presiding Elder") Or ((Role = "Area Head" Or Role = "District Pastor" Or Role = "Presiding Elder") And CurrentlyLeading = "No") Then ' if the member has none of these roles or they do have a role but they arer not leading currently 
                If DatabaseConnection() Then
                    ' Clear_Fields()
                    Dim SQLCMD As New OleDbCommand
                    With SQLCMD
                        .Connection = Connection
                        .CommandText = "Select LOCAL_ID,NUMBER,* From LOCAL_TABLE,MEMBERS_TABLE " & "Where LOCAL_TABLE.LOCAL_ID = MEMBERS_TABLE.LOCALS_ID and MEMBER_ID = @DeleteMemberID "

                        .Parameters.AddWithValue("@DeleteMemberID", Selected_Member_ID)
                        Dim DataReader As OleDbDataReader = .ExecuteReader()
                        While DataReader.Read
                            SubtractNumberInLocalsID = Int(DataReader("LOCAL_ID")) ' assigning the local id and the number in the local that the member was attending 
                            SubtractNumberInLocals = Int(DataReader("NUMBER"))
                        End While
                        DataReader.Close()
                    End With
                    Connection.Close()
                End If
                Delete_Member() ' calls the sub procedure delete member
                CurrentlyLeading = " "
                SubtractFromLocal() ' calling the sub procedure "Subtract from Local"
            Else
                MsgBox(" member you have selected is a leader that has been assined to either a district,local or area. please un assign them from the place they oversee") ' this msg box informs the user the member can not be deleted since they are currently assingned 
            End If
        ElseIf Delete = vbNo Then ' if no is selected then the delete function will be terminated 
            Display_Member_Name_Textbox.Clear()
            Delete_Members_Button.Enabled = False
            CurrentlyLeading = " "
        End If
    End Sub
    Private Sub Load_New_Locals()
        ' this sub procedure loads the locals from the database into the combo box so that the user can select from it
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select LOCAL_NAME From LOCAL_TABLE" ' sql statement to select the locals name

                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim LocalName As String ' local variable will store the names of the local read by the recordset 
                    LocalName = DataReader("LOCAL_NAME")
                    New_Member_Local_ComboBox.Items.Add(LocalName) ' adding the local to the combo box
                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If
    End Sub
    Private Sub Load_Member_New_Roles()
        'this sub procedure will load the roles that a member can have into a combo box so the user can select a role 
        New_Member_Role_ComboBox.Items.Add("Sister")
        New_Member_Role_ComboBox.Items.Add("Brother")
        New_Member_Role_ComboBox.Items.Add("Deacon")
        New_Member_Role_ComboBox.Items.Add("Deaconess")
        New_Member_Role_ComboBox.Items.Add("Presiding Elder")
        New_Member_Role_ComboBox.Items.Add("Elder")
        New_Member_Role_ComboBox.Items.Add("Instrumentalist")
        New_Member_Role_ComboBox.Items.Add("Singer")
        New_Member_Role_ComboBox.Items.Add("District Pastor")
        New_Member_Role_ComboBox.Items.Add("Area Head")
        New_Member_Role_ComboBox.Items.Add("Pastor")
        New_Member_Role_ComboBox.Items.Add("Area Head's Child")
        New_Member_Role_ComboBox.Items.Add("District Pastor's Child")
        New_Member_Role_ComboBox.Items.Add("District Pastor's Wife")
        New_Member_Role_ComboBox.Items.Add("Pastor's Wife")
        New_Member_Role_ComboBox.Items.Add("Pastor's Child")
        New_Member_Role_ComboBox.Items.Add("Area Head's Wife")
        New_Member_Role_ComboBox.Items.Add("Apostle's Wife")
        New_Member_Role_ComboBox.Items.Add("Apostle's Child ")
        New_Member_Role_ComboBox.Items.Add("Apostle")
        New_Member_Role_ComboBox.Items.Add("Youth member")
        New_Member_Role_ComboBox.Items.Add("Youth Leader")
        New_Member_Role_ComboBox.Items.Add("Child")
        New_Member_Role_ComboBox.Items.Add("Sunday School Teacher")
    End Sub
    Private Sub Load_Member_Details_Click(sender As Object, e As EventArgs) Handles Load_Member_Details_Button.Click
        ' when this button is clicked it calls these sub procedures 
        If Member_ID_Listbox.SelectedItem IsNot Nothing Then
            New_Member_Role_ComboBox.Items.Clear() ' this clears the combo boxes
            New_Member_Local_ComboBox.Items.Clear()
            Load_New_Locals()
            Load_Member_New_Roles()
            Load_Member_Details()
            ' calls these sub procedures 
        End If
    End Sub
    Private Sub CheckForLeader()
        'this sub procedure check if the member is currently leading or not and also selects the role of the member 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select * " & "From MEMBERS_TABLE " & "Where MEMBER_ID = @MemberID" ' and LOCAL_TABLE.LOCAL_ID = MEMBERS_TABLE.LOCALS_ID "

                .Parameters.AddWithValue("@MemberID", Selected_Member_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Role = DataReader("MEMBER_ROLE") ' storing the role in the role variable 
                    CurrentlyLeading = DataReader("CURRENTLY_LEADING") ' storing if the member is currently a leader or not 
                End While
                DataReader.Close()
            End With
            MsgBox(Role)
            If Role = "Area Head" Or Role = "Presiding Elder" Or Role = "District Pastor" Then ' if a member has any of these roles
                IsALeader = True ' set the boolean to true 

            ElseIf Role <> "Area Head" And Role <> "Presiding Elder" And Role <> "District Pastor" Then ' if they dont have these roles
                IsALeader = False ' set the boolean to false

            End If
        End If
    End Sub
    Private Sub Load_Member_Details()
        'this sub procedure loads the details of the selected member into the text boxes so their information can be modifyied where te member's ID is the same as the one that was selected by the user
        If DatabaseConnection() Then
            CheckForLeader()
            Dim SQLCMD As New OleDbCommand
            If IsALeader = False And CurrentlyLeading = "No" Then ' if the member is not a leader( presiding elder, Area head or district pastor)
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Select LOCAL_ID,NUMBER,* " & "From LOCAL_TABLE,MEMBERS_TABLE " & "Where MEMBER_ID = @MemberID and LOCAL_TABLE.LOCAL_ID = MEMBERS_TABLE.LOCALS_ID "

                    .Parameters.AddWithValue("@MemberID", Selected_Member_ID)
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read
                        Dim Name, Number, Role, Age As String 'these variables are created in order to hold the data that was read by the recordset 
                        Name = DataReader("MEMBER_NAME")
                        Age = DataReader("MEMBER_AGE")
                        Number = DataReader("MEMBER_TELEPHONE_NUMBER")
                        Role = DataReader("MEMBER_ROLE")
                        Local_ID = Int(DataReader("LOCAL_ID"))

                        New_Member_Name_TextBox.Text = Name ' loading the data into the textboxes
                        New_Member_Age_TextBox.Text = Age
                        New_Member_Telephone_TextBox.Text = Number
                        New_Member_Role_ComboBox.Visible = True
                        New_Member_Local_ComboBox.Visible = True ' setting the combo boxes that display the role and the locals to visable 
                    End While
                    DataReader.Close()
                End With
                Connection.Close() ' closing the connection 
                Update_Member_Button.Enabled = True
            ElseIf IsALeader = True And CurrentlyLeading = "No" Then ' if they are a leader but are not currently leading
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Select * " & "From MEMBERS_TABLE " & "Where MEMBER_ID = @MemberID" ' and LOCAL_TABLE.LOCAL_ID = MEMBERS_TABLE.LOCALS_ID "

                    .Parameters.AddWithValue("@MemberID", Selected_Member_ID)
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read
                        Dim Name, Number, Role, Age As String ' these  local variables store the data that is read from the recordset
                        Name = DataReader("MEMBER_NAME")
                        Age = DataReader("MEMBER_AGE")
                        Number = DataReader("MEMBER_TELEPHONE_NUMBER")
                        Role = DataReader("MEMBER_ROLE")
                        New_Member_Name_TextBox.Text = Name
                        New_Member_Age_TextBox.Text = Age
                        New_Member_Telephone_TextBox.Text = Number
                        New_Member_Role_ComboBox.Visible = True
                        New_Member_Local_ComboBox.Visible = False ' setting the local combo box to false
                    End While
                    DataReader.Close()
                End With
                Connection.Close()
                Update_Member_Button.Enabled = True
            ElseIf CurrentlyLeading = "Yes" Then ' if they are currently leading 
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Select * " & "From MEMBERS_TABLE " & "Where MEMBER_ID = @MemberID" ' and LOCAL_TABLE.LOCAL_ID = MEMBERS_TABLE.LOCALS_ID "

                    .Parameters.AddWithValue("@MemberID", Selected_Member_ID)
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read
                        Dim Name, Number, Age As String ' local variables hold the data read from the recorset 
                        Name = DataReader("MEMBER_NAME")
                        Age = DataReader("MEMBER_AGE")
                        Number = DataReader("MEMBER_TELEPHONE_NUMBER")



                        New_Member_Name_TextBox.Text = Name
                        New_Member_Age_TextBox.Text = Age
                        New_Member_Telephone_TextBox.Text = Number
                        New_Member_Role_ComboBox.Visible = False
                        New_Member_Local_ComboBox.Visible = False ' setting the local and role combo box visability to false 
                    End While
                    DataReader.Close()
                End With
                Connection.Close()
                Update_Member_Button.Enabled = True
            End If
        End If

        ' End If
    End Sub

    Private Sub Update_Member_Button_Click(sender As Object, e As EventArgs) Handles Update_Member_Button.Click
        ' this subprocedure is called when the update button is clicked. this carries out validation checks on the data the user entered
        If New_Member_Name_TextBox.Text = String.Empty Then ' validation checks on the members name
            MsgBox("Please fill in the member's name")
        ElseIf IsNumeric(New_Member_Name_TextBox.Text) Then
            MsgBox("numbers in names are not permitted")
        ElseIf Len(New_Member_Name_TextBox.Text) < 4 Or Len(Member_Name_TextBox.Text) > 15 Then
            MsgBox("Please enter a name that is between 4 and 15 characters")


        ElseIf New_Member_Age_TextBox.Text = String.Empty Then ' validation checks on the members age
            MsgBox("Please fill in the member's age")
        ElseIf Not IsNumeric(New_Member_Age_TextBox.Text) Then
            MsgBox("strings are not permitted in an age")


        ElseIf New_Member_Telephone_TextBox.Text = String.Empty Then ' validation checks on the members phone number
            MsgBox("Please fill in the member's phone number")
        ElseIf Not IsNumeric(New_Member_Telephone_TextBox.Text) Then
            MsgBox("letters are not permitted in a telephone number")
        ElseIf Len(New_Member_Telephone_TextBox.Text) <> 11 Or Int(New_Member_Telephone_TextBox.Text.Substring(0, 1)) <> 0 Then
            MsgBox("Please enter a number that is 11 didgits and 0 is the starting didgit")


        ElseIf New_Member_Role_ComboBox.SelectedItem Is Nothing And New_Member_Role_ComboBox.Visible = True Then ' if a role hasnt been selected and the combo box that displays the role is visable 
            MsgBox("Please select a role for the member")
        ElseIf New_Member_Local_ComboBox.SelectedItem Is Nothing And IsALeader = False And New_Member_Local_ComboBox.Visible = True Then ' if the member is not a leader and they there hasnt been a local chosen 
            MsgBox("Please select a local for the member")
        Else

            Update_Member_Details() ' call this subprocedure


        End If
    End Sub
    Private Sub UpdateNumbersInLocal()
        'this sub procedure is used to determine how the local will be updated, wheter they will loos a member or not when a new member has been updated
        If DatabaseConnection() Then
            If Local_ID <> NewLocal_ID Then ' if the members old local ID is not the same as the members new local ID 
                Dim SQLCMD As New OleDbCommand
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Select NUMBER,LOCAL_ID From LOCAL_TABLE " & "Where LOCAL_TABLE.LOCAL_ID = @LocalID"
                    .Parameters.AddWithValue("@LocalID", Local_ID)
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read

                        SubtractNumberInLocalsID = Int(DataReader("LOCAL_ID"))
                        SubtractNumberInLocals = Int(DataReader("NUMBER"))
                        ' storing the number and the local id that was read 
                    End While
                    DataReader.Close()
                End With
                SubtractFromLocal()
                GetNumberFromLocal() ' calling sub procedure  to get the number in the local
            ElseIf (NewRole = "Area Head" Or NewRole = "District Pastor" Or NewRole = "Presiding Elder") Then ' if the members new role is a presiding elder or area head or district pastor
                Dim SQLCMD As New OleDbCommand
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Select NUMBER,LOCAL_ID From LOCAL_TABLE " & "Where LOCAL_TABLE.LOCAL_ID = @LocalID"
                    .Parameters.AddWithValue("@LocalID", Local_ID)
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read

                        SubtractNumberInLocalsID = Int(DataReader("LOCAL_ID"))
                        SubtractNumberInLocals = Int(DataReader("NUMBER"))
                        ' storing the number and the local id that was read 
                    End While
                    DataReader.Close()
                End With
                SubtractFromLocal() ' calling the sub procedure that reduces the number in locals by one

            End If
        End If


    End Sub
    Private Sub Update_Member_Details()
        'this sub procedure updates the member that corresponds to the member ID that was selected by the user
        If DatabaseConnection() Then
            UpdateNumbersInLocal() ' calls the sub procedure above 
            Dim SQLCMD As New OleDbCommand
            If IsALeader = False And CurrentlyLeading = "No" Then ' if the boolean is false and the member is not leading currently
                With SQLCMD
                    .Connection = Connection

                    .CommandText = "Update MEMBERS_TABLE " & "Set MEMBER_NAME = @MemberName, " & "MEMBER_AGE = @MemberAge, " & "MEMBER_TELEPHONE_NUMBER = @MemberTelNumber, " & "MEMBER_ROLE = @MemberRole, " & "LOCALS_ID = @LocalID " & "Where MEMBER_ID = @MemberID "

                    .Parameters.AddWithValue("@MemberName", New_Member_Name_TextBox.Text)
                    .Parameters.AddWithValue("@MemberAge", New_Member_Age_TextBox.Text)
                    .Parameters.AddWithValue("@MemberTelNumber", New_Member_Telephone_TextBox.Text)
                    .Parameters.AddWithValue("@MemberRole", NewRole)
                    .Parameters.AddWithValue("@LocalID", NewLocal_ID)
                    .Parameters.AddWithValue("@MemberID", Selected_Member_ID)
                    .ExecuteNonQuery()
                    'updates the members details 
                    Update_Member_Button.Enabled = False
                End With
                Connection.Close()
                Display_Members_Table()
                CurrentlyLeading = " "
            ElseIf IsALeader = True And CurrentlyLeading = "No" Then ' if the boolean is true but the member is not  currently leading

                With SQLCMD
                    .Connection = Connection

                    .CommandText = "Update MEMBERS_TABLE " & "Set MEMBER_NAME = @MemberName, " & "MEMBER_AGE = @MemberAge, " & "MEMBER_TELEPHONE_NUMBER = @MemberTelNumber, " & "MEMBER_ROLE = @MemberRole " & "Where MEMBER_ID = @MemberID "

                    .Parameters.AddWithValue("@MemberName", New_Member_Name_TextBox.Text)
                    .Parameters.AddWithValue("@MemberAge", New_Member_Age_TextBox.Text)
                    .Parameters.AddWithValue("@MemberTelNumber", New_Member_Telephone_TextBox.Text)
                    .Parameters.AddWithValue("@MemberRole", NewRole)
                    .Parameters.AddWithValue("@MemberID", Selected_Member_ID)
                    .ExecuteNonQuery()
                    'update the members details 
                    Update_Member_Button.Enabled = False
                End With

                Connection.Close()
                Display_Members_Table()
                CurrentlyLeading = " "
            ElseIf CurrentlyLeading = "Yes" Then ' if the member is currently leading 
                With SQLCMD
                    .Connection = Connection

                    .CommandText = "Update MEMBERS_TABLE " & "Set MEMBER_NAME = @MemberName, " & "MEMBER_AGE = @MemberAge, " & "MEMBER_TELEPHONE_NUMBER = @MemberTelNumber " & "Where MEMBER_ID = @MemberID "
                    .Parameters.AddWithValue("@MemberName", New_Member_Name_TextBox.Text)
                    .Parameters.AddWithValue("@MemberAge", New_Member_Age_TextBox.Text)
                    .Parameters.AddWithValue("@MemberTelNumber", New_Member_Telephone_TextBox.Text)
                    .Parameters.AddWithValue("@MemberID", Selected_Member_ID)
                    .ExecuteNonQuery() ' updating the members details 
                    Update_Member_Button.Enabled = False
                End With
                Connection.Close()
                Display_Members_Table()
                CurrentlyLeading = " "
            End If
        End If
    End Sub

    Private Sub New_Member_Role_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles New_Member_Role_ComboBox.SelectedIndexChanged
        'this sub procedure stores the value of the new role for the member that has been selected by the user 
        If New_Member_Role_ComboBox.SelectedItem IsNot Nothing Then ' checks if something has been selected 
            NewRole = New_Member_Role_ComboBox.SelectedItem.ToString
            If NewRole = "Area Head" Or NewRole = "Presiding Elder" Or NewRole = "District Pastor" Then
                IsALeader = True ' boolean is true 
                New_Member_Local_ComboBox.Visible = False ' local combobox is invisable 
            ElseIf NewRole <> "Area Head" And NewRole <> "Presiding Elder" And NewRole <> "District Pastor" Then
                IsALeader = False ' boolean is false 
                New_Member_Local_ComboBox.Visible = True ' local combobox is invisable 
            End If
        End If
    End Sub

    Private Sub New_Member_Local_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles New_Member_Local_ComboBox.SelectedIndexChanged
        ' this sub procedure stores the value of the new local for the member that has been selected by the user 
        Dim LocalName As String ' local variable that will hold the name of the local that was selected 
        If New_Member_Local_ComboBox.SelectedItem IsNot Nothing Then
            LocalName = New_Member_Local_ComboBox.SelectedItem.ToString 'storing the local selected from the combo box
            If DatabaseConnection() Then
                Dim SQLCMD As New OleDbCommand
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Select LOCAL_ID,NUMBER From LOCAL_TABLE " & "Where LOCAL_TABLE.LOCAL_NAME = @LocalName"
                    .Parameters.AddWithValue("@LocalName", LocalName)
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read
                        NewLocal_ID = Int(DataReader("LOCAL_ID")) 'storing the new local ID
                    End While
                    DataReader.Close()
                End With
                Connection.Close() 'close the connection 
            End If

        End If
    End Sub

    Private Sub Search_Member_Name_Button_Click(sender As Object, e As EventArgs) Handles Search_Member_Name_Button.Click
        'this sub procedure is called when the search by name button is clicked. it carries out validation checks on the name entered by the user
        Dim Member_Name As String ' variable will hold the name that the user types into the search textbox 
        If Search_Member_Name_TextBox.Text = String.Empty Then ' validation checks on the name
            MsgBox("Please fill in the member's name")
        ElseIf IsNumeric(Member_Name_TextBox.Text) Then
            MsgBox("numbers in names are not permitted")
        Else
            Member_Name = Search_Member_Name_TextBox.Text
            Search_Member_Name(Member_Name) ' calls the sub procedure 
        End If
    End Sub


    Private Sub Search_Member_Name(MemberName As String)
        ' this sub procedure searches the database for a member that matches the name entered or a similar name
        If DatabaseConnection() Then
            Clear_Fields() ' clears everything
            Dim SQLcmd As New OleDbCommand
            With SQLcmd
                .Connection = Connection
                .CommandText = "Select LOCAL_NAME,* " & "From LOCAL_TABLE,MEMBERS_TABLE " & "Where LOCAL_TABLE.LOCAL_ID = MEMBERS_TABLE.LOCALS_ID and MEMBER_NAME Like @NameMember " ' and LOCAL_TABLE.LOCAL_ID = MEMBERS_TABLE.LOCAL.ID " 

                .Parameters.AddWithValue("@NameMember", "%" & MemberName & "%") 'if the name entered is simillar to one in the database 
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim MemberID, Name, Number, Role, Age, Local As String ' these variables store the data read by the record set 
                    MemberID = DataReader("MEMBER_ID")
                    Name = DataReader("MEMBER_NAME")
                    Age = DataReader("MEMBER_AGE")
                    Number = DataReader("MEMBER_TELEPHONE_NUMBER")
                    Role = DataReader("MEMBER_ROLE")
                    Local = DataReader("LOCAL_NAME")
                    Member_ID_Listbox.Items.Add(MemberID)
                    Member_Name_Listbox.Items.Add(Name)
                    Member_Age_Listbox.Items.Add(Age)
                    Member_Number_Listbox.Items.Add(Number)
                    Member_Role_Listbox.Items.Add(Role)
                    Member_Local_Listbox.Items.Add(Local)
                End While
                DataReader.Close()
            End With
            If Member_ID_Listbox.Items.Count = 0 Then 'if there is nothing in the list box select members who dont have a local ID
                With SQLcmd
                    .Connection = Connection
                    .CommandText = "Select * " & "From MEMBERS_TABLE " & "Where MEMBER_NAME Like @NameMember " '

                    .Parameters.AddWithValue("@NameMember", "%" & MemberName & "%") ' if the name entered is simillar to one in the database 
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read
                        Dim MemberID, Name, Number, Role, Age As String ' stores the data that has been read by the recordset
                        MemberID = DataReader("MEMBER_ID")
                        Name = DataReader("MEMBER_NAME")
                        Age = DataReader("MEMBER_AGE")
                        Number = DataReader("MEMBER_TELEPHONE_NUMBER")
                        Role = DataReader("MEMBER_ROLE")

                        Member_ID_Listbox.Items.Add(MemberID)
                        Member_Name_Listbox.Items.Add(Name)
                        Member_Age_Listbox.Items.Add(Age)
                        Member_Number_Listbox.Items.Add(Number)
                        Member_Role_Listbox.Items.Add(Role)

                    End While
                    DataReader.Close()
                End With
            End If
            Connection.Close()
        End If
    End Sub

    Private Sub Search_Member_ID_Button_Click(sender As Object, e As EventArgs) Handles Search_Member_ID_Button.Click
        'this sub procedure is called when the search by ID button is clicked. it carries out validation checks on the ID entered by the user
        Dim ID As Integer 'storing the ID entered by the user

        If Search_Member_ID_TextBox.Text = String.Empty Then ' validation checks 
            MsgBox("Please fill in the member's ID")
        ElseIf Not IsNumeric(Search_Member_ID_TextBox.Text) Then
            MsgBox("strings are not permitted in an ID")
        Else
            ID = Int(Search_Member_ID_TextBox.Text)
            Search_Member_ID(ID) ' passes the parameter through
        End If
    End Sub

    Private Sub Search_Member_ID(Member_ID As Integer) ' the passed parameter 
        'this sub procedure searches the database for the record that corresponds to the id that the user typed in the text box 
        If DatabaseConnection() Then
            Clear_Fields() ' calls the clear fields sub procedure 
            Dim SQLcmd As New OleDbCommand
            With SQLcmd
                .Connection = Connection
                .CommandText = "Select LOCAL_NAME,* " & "From LOCAL_TABLE,MEMBERS_TABLE " & "Where LOCAL_TABLE.LOCAL_ID = MEMBERS_TABLE.LOCALS_ID and MEMBER_ID = @MemberID " 'This is very similar to the StudentID Select query but this time we’re searching using the Surname field and the recordset returned can contain multiple records

                .Parameters.AddWithValue("@MemberID", Member_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim MemberID, Name, Number, Role, Age, Local As String 'stores the info read by the recordset 
                    MemberID = DataReader("MEMBER_ID")
                    Name = DataReader("MEMBER_NAME")
                    Age = DataReader("MEMBER_AGE")
                    Number = DataReader("MEMBER_TELEPHONE_NUMBER")
                    Role = DataReader("MEMBER_ROLE")
                    Local = DataReader("LOCAL_NAME")
                    Member_ID_Listbox.Items.Add(MemberID)
                    Member_Name_Listbox.Items.Add(Name)
                    Member_Age_Listbox.Items.Add(Age)
                    Member_Number_Listbox.Items.Add(Number)
                    Member_Role_Listbox.Items.Add(Role)
                    Member_Local_Listbox.Items.Add(Local)



                End While
                DataReader.Close()
            End With
            If Member_ID_Listbox.Items.Count = 0 Then ' if there is nothing displayed in the list boxes 
                With SQLcmd
                    .Connection = Connection
                    .CommandText = "Select * " & "From MEMBERS_TABLE " & "Where MEMBER_ID Like @MemberID "

                    .Parameters.AddWithValue("@MemberID", Member_ID)
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read
                        Dim MemberID, Name, Number, Role, Age As String 'stores the info read by the recordset
                        MemberID = DataReader("MEMBER_ID")
                        Name = DataReader("MEMBER_NAME")
                        Age = DataReader("MEMBER_AGE")
                        Number = DataReader("MEMBER_TELEPHONE_NUMBER")
                        Role = DataReader("MEMBER_ROLE")

                        Member_ID_Listbox.Items.Add(MemberID)
                        Member_Name_Listbox.Items.Add(Name)
                        Member_Age_Listbox.Items.Add(Age)
                        Member_Number_Listbox.Items.Add(Number)
                        Member_Role_Listbox.Items.Add(Role) 'displaying it in the list boxes 

                    End While
                    DataReader.Close()
                End With
            End If
            Connection.Close() ' close the connection
        End If
    End Sub

    Private Sub Add_Member_Clear_Button_Click(sender As Object, e As EventArgs) Handles Add_Member_Clear_Button.Click
        ' clears the user inputs
        Member_Name_TextBox.Clear()
        Member_Age_TextBox.Clear()
        Member_Telephone_TextBox.Clear()
        Add_Member_Local_ComboBox.SelectedIndex = -1
        Add_Member_Role_ComboBox.SelectedIndex = -1
    End Sub

    Private Sub Search_Member_Clear_Button_Click(sender As Object, e As EventArgs) Handles Search_Member_Clear_Button.Click
        ' clears the user inputs
        Search_Member_ID_TextBox.Clear()
        Search_Member_Name_TextBox.Clear()
    End Sub

    Private Sub Update_Member_Clear_Button_Click(sender As Object, e As EventArgs) Handles Update_Member_Clear_Button.Click
        ' clears the user inputs
        New_Member_Name_TextBox.Clear()
        New_Member_Age_TextBox.Clear()
        New_Member_Telephone_TextBox.Clear()
        New_Member_Local_ComboBox.SelectedIndex = -1
        New_Member_Role_ComboBox.SelectedIndex = -1
    End Sub


End Class