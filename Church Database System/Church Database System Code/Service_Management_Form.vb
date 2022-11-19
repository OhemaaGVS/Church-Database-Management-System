Imports System.Data.OleDb
Public Class Service_Management_Form
    Private ValueOfServiceID As Integer = -1 ' this variable will store the value of the services id  
    Private Local_ID As Integer ' this variable will store the local id that has been assigned to the service 
    Private Service_Type As String ' this variable will store the type of service the service is 
    Private Selected_Service_ID As Integer ' this will store the service id that has been selected from the list box by the user
    Private New_Service_Type As String ' this variable will store the services new type 
    Private NewLocal_ID As Integer 'This variable will store the new local id for the service
    Private Sub Service_ID_Listbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Service_ID_Listbox.SelectedIndexChanged
        ' this sub procedure stores the service selected by the user 
        If Service_ID_Listbox.SelectedItem IsNot Nothing Then ' checks  if ther has been an id selected 
            Selected_Service_ID = Int(Service_ID_Listbox.SelectedItem) ' if there has then store it
        End If
    End Sub

    Private Sub Add_Service_Button_Click(sender As Object, e As EventArgs) Handles Add_Service_Button.Click
        ' when this button is clicked it carries out all the validations needed on the inputs
        If Add_Service_Name_TextBox.Text = String.Empty Then ' Validation checks on the services name 
            MsgBox("Please fill in the service's name")
        ElseIf IsNumeric(Add_Service_Name_TextBox.Text) Then
            MsgBox("numbers in names are not permitted")
        ElseIf Len(Add_Service_Name_TextBox.Text) < 4 Or Len(Add_Service_Name_TextBox.Text) > 25 Then
            MsgBox("Please enter a name that is between 4 and 25 characters")


        ElseIf Add_Service_Attendance_TextBox.Text = String.Empty Then 'Validation checks on the services attendance
            MsgBox("Please fill in the servics attendance")
        ElseIf Not IsNumeric(Add_Service_Attendance_TextBox.Text) Then
            MsgBox("strings are not permitted")


        ElseIf Service_Type_ComboBox.SelectedItem Is Nothing Then 'Validation to check if the service type has been selected
            MsgBox("Please select a servuce type")
        ElseIf Service_Local_ComboBox.SelectedItem Is Nothing Then 'Validation to check if a local has been chosen 
            MsgBox("Please select a local")
        Else
            Create_New_Service()
        End If

    End Sub

    Private Sub Create_New_Service()
        ' this sub procedure saves the new service created into the database 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            If ValueOfServiceID = -1 Then
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Insert into SERVICE_TABLE (SERVICE_NAME, SERVICE_TYPE, SERVICE_DATE, SERVICE_ATTENDANCE, LOCAL_ID )" & "Values (@ServiceName ,@ServiceType ,@ServiceDate ,@ServiceAttendance ,@LocalID )" 'sql insertion 
                    .Parameters.AddWithValue("@ServiceName", Add_Service_Name_TextBox.Text)
                    .Parameters.AddWithValue("@ServiceType", Service_Type)
                    .Parameters.AddWithValue("@ServiceDate ", Service_Date_DatePicker.Value.ToLongDateString)
                    .Parameters.AddWithValue("@ServiceAttendance", Add_Service_Attendance_TextBox.Text)
                    .Parameters.AddWithValue("@LocalID", Local_ID)
                    .ExecuteNonQuery()
                    .CommandText = "Select @@Identity"
                    ValueOfServiceID = .ExecuteScalar
                    Add_Service_Auto_ID_Label.Text = ValueOfServiceID
                End With
                Connection.Close()
                Add_Service_Auto_ID_Label.Text = "Automatically Generated"
                Display_Services_Table() ' calls the dsiplay services function
            End If

        End If

    End Sub

    Private Sub Clear_Fields()
        ' this sub procedure clears the textboxes and the other user inputs 
        Service_ID_Listbox.Items.Clear()
        Service_Name_Listbox.Items.Clear()
        Service_Type_Listbox.Items.Clear()
        Service_Date_Listbox.Items.Clear()
        Service_Attendance_Listbox.Items.Clear()
        Service_Local_ListBox.Items.Clear()
        Add_Service_Name_TextBox.Clear()
        Search_For_Service_Name_TextBox.Clear()
        Search_For_Service_ID_TextBox.Clear()
        Display_Service_Name_Textbox.Clear()
        Service_Date_DatePicker.Value = Now()
        New_Service_Date_DatePicker.Value = Now()
        Add_Service_Attendance_TextBox.Clear()
        New_Service_Name_TextBox.Clear()
        New_Service_Attendance_TextBox.Clear()
        Add_Service_Attendance_TextBox.Clear()
        Service_Local_ComboBox.SelectedIndex = -1
        Service_Type_ComboBox.SelectedIndex = -1
        New_Service_Type_ComboBox.SelectedIndex = -1
        New_Service_Local_ComboBox.SelectedIndex = -1
        New_Service_Type_ComboBox.Items.Clear()
        New_Service_Local_ComboBox.Items.Clear()
        Service_Local_ComboBox.Items.Clear()
        Service_Type_ComboBox.Items.Clear()
        ValueOfServiceID = -1

    End Sub

    Private Sub Display_Services_Table()
        ' this sub procedure displays the services and its information
        If DatabaseConnection() Then
            Clear_Fields()
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select LOCAL_NAME,* From LOCAL_TABLE,SERVICE_TABLE " & "Where LOCAL_TABLE.LOCAL_ID = SERVICE_TABLE.LOCAL_ID"

                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim ServiceID, ServiceName, Type, ServiceDate, Attendance, Local As String ' these variables will store the data that is read by the recordset 
                    ServiceID = DataReader("SERVICE_ID")
                    ServiceName = DataReader("SERVICE_NAME")
                    Type = DataReader("SERVICE_TYPE")
                    ServiceDate = DataReader("SERVICE_DATE")
                    Attendance = DataReader("SERVICE_ATTENDANCE")
                    Local = DataReader("LOCAL_NAME")
                    Service_ID_Listbox.Items.Add(ServiceID)
                    Service_Name_Listbox.Items.Add(ServiceName)
                    Service_Type_Listbox.Items.Add(Type)
                    Service_Date_Listbox.Items.Add(ServiceDate)
                    Service_Attendance_Listbox.Items.Add(Attendance)
                    Service_Local_ListBox.Items.Add(Local) ' displaying the data
                End While
                DataReader.Close()
            End With
            Connection.Close() ' closing the connection
        End If
        Load_Service_Types()
        Load_Locals()
    End Sub

    Private Sub Service_Management_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' this sub procedure is called when the form loads, it calls several procedures and disables some  buttons
        Display_Services_Table()
        Delete_Service_Button.Enabled = False
        Update_Service_Details_Button.Enabled = False ' disabling buttons
    End Sub
    Private Sub Load_Service_Types()
        ' this sub procedure loads all of the different service types into a combo box so it can be seleceted by the user
        Service_Type_ComboBox.Items.Add("All Night")
        Service_Type_ComboBox.Items.Add("Prayer Meeting")
        Service_Type_ComboBox.Items.Add("Half Night")
        Service_Type_ComboBox.Items.Add("Thanks Giving")
        Service_Type_ComboBox.Items.Add("Baptsim ")
        Service_Type_ComboBox.Items.Add("Naming")
        Service_Type_ComboBox.Items.Add("Welcome Service")
        Service_Type_ComboBox.Items.Add("Farewell Service")
        Service_Type_ComboBox.Items.Add("Wedding")
        Service_Type_ComboBox.Items.Add("Engagment Ceremony")
        Service_Type_ComboBox.Items.Add("Funeral")
    End Sub

    Private Sub Service_Type_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Service_Type_ComboBox.SelectedIndexChanged
        ' this sub procedure stores the type of service the user has selected from the combo box 
        If Service_Type_ComboBox.SelectedItem IsNot Nothing Then ' if a service type has been selected 
            Service_Type = Service_Type_ComboBox.SelectedItem.ToString ' stores the service type 
        End If
    End Sub

    Private Sub Service_Local_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Service_Local_ComboBox.SelectedIndexChanged
        ' this subprocedure will store the local id that corrosponds to the local name chosen by the user 
        Dim LocalName As String 'this variable will store the name of the local chosen from the combobox by the user
        If Service_Local_ComboBox.SelectedItem IsNot Nothing Then
            LocalName = Service_Local_ComboBox.SelectedItem.ToString ' storing the name
            If DatabaseConnection() Then
                Dim SQLCMD As New OleDbCommand
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Select LOCAL_ID From LOCAL_TABLE " & "Where LOCAL_TABLE.LOCAL_NAME = @LocalName"
                    .Parameters.AddWithValue("@LocalName", LocalName)
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read
                        Local_ID = Int(DataReader("LOCAL_ID")) ' storing the local id
                    End While
                    DataReader.Close()
                End With
                Connection.Close()
            End If

        End If
    End Sub
    Private Sub Load_Locals()
        ' this sub procedure loads all of the locals into the combo box so that it can be selected by the user 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select LOCAL_NAME From LOCAL_TABLE" ' selecting the locals

                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim Local As String ' this will hold the name read by the recordset
                    Local = DataReader("LOCAL_NAME")
                    Service_Local_ComboBox.Items.Add(Local) 'adding it into the combo box
                End While
                DataReader.Close()
            End With
            Connection.Close() ' connection 

        End If
    End Sub

    Private Sub Load_Service_Name_Button_Click(sender As Object, e As EventArgs) Handles Load_Service_Name_Button.Click
        ' this sub procedure calls the sub procedure load service name 
        If Service_ID_Listbox.SelectedItem IsNot Nothing Then
            Load_Service_Name()
        End If


    End Sub
    Private Sub Load_Service_Name()
        ' this procedure loads the services name into the textbox, if the service's if its id is the same as the one selected by the user 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select SERVICE_NAME From SERVICE_TABLE " & "Where SERVICE_ID = @ServiceID"
                .Parameters.AddWithValue("@ServiceID", Selected_Service_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim ServiceName As String ' this will hold the name that was read from the record set
                    ServiceName = DataReader("SERVICE_NAME")
                    Display_Service_Name_Textbox.Text = ServiceName 'display the services name 
                End While
                DataReader.Close()
            End With
            Connection.Close() ' close the connection
        End If
        Delete_Service_Button.Enabled = True
    End Sub
    Private Sub Delete_Service()
        ' this is sub procedure deletes the service from the database, where the service id is the same as the id that was selected by the user  
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Delete * " & "From SERVICE_TABLE " & "Where SERVICE_ID = @DeleteServiceID " ' sql delete command 
                .Parameters.AddWithValue("@DeleteServiceID", Selected_Service_ID)

                .ExecuteNonQuery()
            End With
            Connection.Close() ' close the connection 
            Display_Services_Table()
            Display_Service_Name_Textbox.Clear()
            Delete_Service_Button.Enabled = False
        End If
    End Sub

    Private Sub Delete_Service_Click(sender As Object, e As EventArgs) Handles Delete_Service_Button.Click
        'this sub verifies if the user is sure they want to delete the service
        Dim Delete As String ' 'creating the local variable delete
        Delete = MsgBox("Are you sure you would like to delete this Service? Data will be permenantly deleted", vbExclamation + vbYesNo + vbDefaultButton2, "Delete Service Confirmation")

        If Delete = vbYes Then ' if yes is selected
            Delete_Service()

        ElseIf Delete = vbNo Then ' if no is selected
            Display_Service_Name_Textbox.Clear()
            Delete_Service_Button.Enabled = False
        End If
    End Sub
    Private Sub Load_New_Locals()
        ' this sub procedure loads the names of the locals into a combo box
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select LOCAL_NAME From LOCAL_TABLE"

                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim Local As String '  this will hold the data read by the record set
                    Local = DataReader("LOCAL_NAME")
                    New_Service_Local_ComboBox.Items.Add(Local) ' adding them to a combo box
                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If
    End Sub
    Private Sub Load_New_Service_Types()
        ' this sub loads the different types of services into a combo box
        New_Service_Type_ComboBox.Items.Add("All Night")
        New_Service_Type_ComboBox.Items.Add("Prayer Meeting")
        New_Service_Type_ComboBox.Items.Add("Half Night")
        New_Service_Type_ComboBox.Items.Add("Thanks Giving")
        New_Service_Type_ComboBox.Items.Add("Baptsim ")
        New_Service_Type_ComboBox.Items.Add("Naming")
        New_Service_Type_ComboBox.Items.Add("Welcome Service")
        New_Service_Type_ComboBox.Items.Add("Farewell Service")
        New_Service_Type_ComboBox.Items.Add("Wedding")
        New_Service_Type_ComboBox.Items.Add("Engagment Ceremony")
        New_Service_Type_ComboBox.Items.Add("Funeral")
    End Sub
    Private Sub Load_Service_Details_Button_Click(sender As Object, e As EventArgs) Handles Load_Service_Details_Button.Click

        ' when this button is clicked it then calls these sub procedures 
        If Service_ID_Listbox.SelectedItem IsNot Nothing Then
            New_Service_Type_ComboBox.Items.Clear()
            New_Service_Local_ComboBox.Items.Clear()
            Load_New_Locals()
            Load_New_Service_Types()
            Load_Services_Details()
        End If
    End Sub
    Private Sub Load_Services_Details()
        ' this sub procedure loads the current details about the service into the text boxes so that they can be ammended to
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select LOCAL_NAME,* " & "From LOCAL_TABLE,SERVICE_TABLE " & "Where SERVICE_ID = @ServiceID and LOCAL_TABLE.LOCAL_ID = SERVICE_TABLE.LOCAL_ID "

                .Parameters.AddWithValue("@ServiceID", Selected_Service_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim ServiceName, Type, ServiceDate, Attendance, Local As String ' these variables will hold the data read by the record set
                    ServiceName = DataReader("SERVICE_NAME")
                    Type = DataReader("SERVICE_TYPE")
                    ServiceDate = DataReader("SERVICE_DATE")
                    Attendance = DataReader("SERVICE_ATTENDANCE")
                    Local = DataReader("LOCAL_NAME")
                    New_Service_Name_TextBox.Text = ServiceName
                    New_Service_Attendance_TextBox.Text = Attendance ' display them


                End While
                DataReader.Close() ' close the record set 
            End With
            Connection.Close() ' close the connection 
            Update_Service_Details_Button.Enabled = True
        End If
    End Sub

    Private Sub Update_Service_Details_Button_Click(sender As Object, e As EventArgs) Handles Update_Service_Details_Button.Click
        ' this sub carries out the validation checks on the users new input 
        If New_Service_Name_TextBox.Text = String.Empty Then ' Validation checks on the services name 
            MsgBox("Please fill in the service's name")
        ElseIf IsNumeric(New_Service_Name_TextBox.Text) Then
            MsgBox("numbers in names are not permitted")
        ElseIf Len(New_Service_Name_TextBox.Text) < 4 Or Len(New_Service_Name_TextBox.Text) > 25 Then
            MsgBox("Please enter a name that is between 4 and 25 characters")


        ElseIf New_Service_Attendance_TextBox.Text = String.Empty Then 'Validation checks on the services attendance
            MsgBox("Please fill in the servics attendance")
        ElseIf Not IsNumeric(New_Service_Attendance_TextBox.Text) Then
            MsgBox("strings are not permitted")


        ElseIf New_Service_Type_ComboBox.SelectedItem Is Nothing Then 'Validation to check if the service type has been selected
            MsgBox("Please select a servuce type")
        ElseIf New_Service_Local_ComboBox.SelectedItem Is Nothing Then 'Validation to check if a local has been chosen 
            MsgBox("Please select a local")
        Else
            Update_Service_Details()
        End If
    End Sub

    Private Sub Update_Service_Details()
        ' this sub procedure updates the services details that corrospong to the id that was selected by the user 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Update SERVICE_TABLE " & "Set SERVICE_NAME = @ServiceName, " & "SERVICE_TYPE = @Type, " & "SERVICE_DATE = @Date, " & "SERVICE_ATTENDANCE = @Attendance, " & "LOCAL_ID = @LocalID " & "Where SERVICE_ID = @ServiceID "

                .Parameters.AddWithValue("@ServiceName", New_Service_Name_TextBox.Text)
                .Parameters.AddWithValue("@Type", New_Service_Type)
                .Parameters.AddWithValue(" @Date", New_Service_Date_DatePicker.Value.ToLongDateString)
                .Parameters.AddWithValue("@Attendance", New_Service_Attendance_TextBox.Text)
                .Parameters.AddWithValue("@LocalID", NewLocal_ID)
                .Parameters.AddWithValue("@MemberID", Selected_Service_ID)
                ' updataing the record 
                .ExecuteNonQuery()
                Update_Service_Details_Button.Enabled = False
            End With
            Connection.Close() ' close the connection
            Display_Services_Table() ' display the details 
        End If
    End Sub

    Private Sub New_Service_Type_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles New_Service_Type_ComboBox.SelectedIndexChanged
        ' this sub procedure is used to store the new service type that the user has selected 
        If New_Service_Type_ComboBox.SelectedItem IsNot Nothing Then
            New_Service_Type = New_Service_Type_ComboBox.SelectedItem.ToString ' assigning it to the varaible

        End If
    End Sub

    Private Sub New_Service_Local_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles New_Service_Local_ComboBox.SelectedIndexChanged
        ' this sub procedure selects the new local id that the service has been given
        Dim LocalName As String
        If New_Service_Local_ComboBox.SelectedItem IsNot Nothing Then
            LocalName = New_Service_Local_ComboBox.SelectedItem.ToString ' storing the name that the user selected 
            If DatabaseConnection() Then
                Dim SQLCMD As New OleDbCommand
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Select LOCAL_ID From LOCAL_TABLE " & "Where LOCAL_TABLE.LOCAL_NAME = @LocalName"
                    .Parameters.AddWithValue("@LocalName", LocalName)
                    Dim DataReader As OleDbDataReader = .ExecuteReader()
                    While DataReader.Read
                        NewLocal_ID = Int(DataReader("LOCAL_ID")) ' storing the local id that corresponds to the local the user chose 
                    End While
                    DataReader.Close()
                End With
                Connection.Close() ' close the connection
            End If

        End If
    End Sub

    Private Sub Search_Servcie_Name_Button_Click(sender As Object, e As EventArgs) Handles Search_Servcie_Name_Button.Click
        ' this sub validates the input of the user and calls the sub that searches for the service
        Dim Name As String ' stores the name inputted by the user 
        If Search_For_Service_Name_TextBox.Text = String.Empty Then ' validation checks on the name
            MsgBox("Please fill in the services name")
        ElseIf IsNumeric(Search_For_Service_Name_TextBox.Text) Then
            MsgBox("numbers in names are not permitted")
        Else
            Name = Search_For_Service_Name_TextBox.Text
            Search_Service_Name(Name) ' calls the sub procedure 
        End If
        
    End Sub


    Private Sub Search_Service_Name(Name As String)
        ' this sub procedure searches for a service within the database 
        If DatabaseConnection() Then
            Clear_Fields()
            Dim SQLcmd As New OleDbCommand
            With SQLcmd
                .Connection = Connection
                .CommandText = "Select LOCAL_NAME,* " & "From LOCAL_TABLE,SERVICE_TABLE " & "Where LOCAL_TABLE.LOCAL_ID = SERVICE_TABLE.LOCAL_ID and SERVICE_NAME Like @Service "

                .Parameters.AddWithValue("@Service", "%" & Name & "%") ' if a similar name is entered 
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim ServiceID, ServiceName, Type, ServiceDate, Attendance, Local As String ' these will hold the values read by the record set 
                    ServiceID = DataReader("SERVICE_ID")
                    ServiceName = DataReader("SERVICE_NAME")
                    Type = DataReader("SERVICE_TYPE")
                    ServiceDate = DataReader("SERVICE_DATE")
                    Attendance = DataReader("SERVICE_ATTENDANCE")
                    Local = DataReader("LOCAL_NAME")
                    Service_ID_Listbox.Items.Add(ServiceID)
                    Service_Name_Listbox.Items.Add(ServiceName)
                    Service_Type_Listbox.Items.Add(Type)
                    Service_Date_Listbox.Items.Add(ServiceDate)
                    Service_Attendance_Listbox.Items.Add(Attendance)
                    Service_Local_ListBox.Items.Add(Local) ' output the results
                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If
    End Sub

    Private Sub Search_Service_ID_Button_Click(sender As Object, e As EventArgs) Handles Search_Service_ID_Button.Click
        'validates the user input into the search id textbox
        Dim ID As Integer
        If Search_For_Service_ID_TextBox.Text = String.Empty Then ' validation checks 
            MsgBox("Please fill in the service's ID")
        ElseIf Not IsNumeric(Search_For_Service_ID_TextBox.Text) Then
            MsgBox("strings are not permitted in an ID")
        Else
            ID = Int(Search_For_Service_ID_TextBox.Text)
            Search_Service_ID(ID) 'passes the parameter through
        End If
       
    End Sub

    Private Sub Search_Service_ID(Service_ID As Integer) ' the passed in parameter
        'searches for record with matching id
        If DatabaseConnection() Then
            Clear_Fields() ' clear the fields 
            Dim SQLcmd As New OleDbCommand
            With SQLcmd
                .Connection = Connection
                .CommandText = "Select LOCAL_NAME,* " & "From LOCAL_TABLE,SERVICE_TABLE " & "Where LOCAL_TABLE.LOCAL_ID = SERVICE_TABLE.LOCAL_ID and SERVICE_ID = @ServiceID "

                .Parameters.AddWithValue("@ServiceID", Service_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim ServiceID, Name, Type, ServiceDate, Attendance, Local As String
                    ServiceID = DataReader("SERVICE_ID")
                    Name = DataReader("SERVICE_NAME")
                    Type = DataReader("SERVICE_TYPE")
                    ServiceDate = DataReader("SERVICE_DATE")
                    Attendance = DataReader("SERVICE_ATTENDANCE")
                    Local = DataReader("LOCAL_NAME")
                    Service_ID_Listbox.Items.Add(ServiceID)
                    Service_Name_Listbox.Items.Add(Name)
                    Service_Type_Listbox.Items.Add(Type)
                    Service_Date_Listbox.Items.Add(ServiceDate)
                    Service_Attendance_Listbox.Items.Add(Attendance)
                    Service_Local_ListBox.Items.Add(Local)
                    ' display the records 
                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If
    End Sub

    Private Sub Service_Clear_Button_Click(sender As Object, e As EventArgs) Handles Service_Clear_Button.Click
        ' clears the inputs 
        Add_Service_Name_TextBox.Clear()
        Add_Service_Attendance_TextBox.Clear()
        Service_Date_DatePicker.Value = Now()
        Service_Local_ComboBox.SelectedIndex = -1
        Service_Type_ComboBox.SelectedIndex = -1
    End Sub

    Private Sub Update_Service_Clear_Button_Click(sender As Object, e As EventArgs) Handles Update_Service_Clear_Button.Click
        ' clears the inputs 
        New_Service_Name_TextBox.Clear()
        New_Service_Attendance_TextBox.Clear()
        New_Service_Local_ComboBox.SelectedIndex = -1
        New_Service_Type_ComboBox.SelectedIndex = -1
        New_Service_Date_DatePicker.Value = Now()
    End Sub

    Private Sub Search_Service_Clear_Button_Click(sender As Object, e As EventArgs) Handles Search_Service_Clear_Button.Click
        ' clears the inputs 
        Search_For_Service_Name_TextBox.Clear()
        Search_For_Service_ID_TextBox.Clear()
    End Sub
End Class