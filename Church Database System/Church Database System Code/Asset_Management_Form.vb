Imports System.Data.OleDb
Public Class Asset_Management_Form
    Private ValueOfAssetID As Integer = -1 ' this variable will store the id given to an asset when it is created 
    Private ValueOfLocal_AssetID As Integer = -1 'this variable will store the id given to an assest assigned to a local when its being created 
    Private Local_ID As Integer '' this variable will store the local if that the asset has been assined to 
    Private Asset_Type, Selected_Asset_Local As String ' asset type will store the type of asset the asset is. selected asset local :this varaible will store the local name selected by the user from this listbox
    Private Selected_Asset_ID, Asset_Price, Asset_Quantity, New_Asset_Price, New_Quantity, Asset_Local_ID As Integer
    'selected asset id: this will store the id that is selected from the list box of assets
    'asset price: this variable will store the price of an asset
    'asset quantity: this will store the quantity of an asset
    'new asset price: this will store the new price of the asset
    ' new asset quantity will store the new quantity of the asset 
    'asset local id will store the id of the local selected by the user in the list box
    Private New_Asset_Type As String ' this variable will store the new type for the asset 
    Private Sub Asset_ID_Listbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Asset_ID_Listbox.SelectedIndexChanged
        ' this sub procedure stores the asset id selected by the user 
        If Asset_ID_Listbox.SelectedItem IsNot Nothing Then
            Selected_Asset_ID = Int(Asset_ID_Listbox.SelectedItem) ' storing the id
            MsgBox(Selected_Asset_ID)
            Fill_TextBoxes()
        End If
    End Sub
    Private Sub Fill_TextBoxes()
        ' this loads the asset name that matches the selected asset id into a textbox
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select ASSET_NAME " & "From ASSET_TABLE " & "Where ASSET_ID = @AssetID" ' where the id is the same as the on selected
                .Parameters.AddWithValue("@AssetID", Selected_Asset_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim Name As String 'this will store the name read by the recorset
                    Name = DataReader("ASSET_NAME")
                    Assign_Asset_Name_TextBox.Text = Name
                End While
            End With
        End If
        Load_Locals()
    End Sub
    Private Sub Add_Asset_Button_Click(sender As Object, e As EventArgs) Handles Add_Asset_Button.Click
        ' when this button is clicked it carries out all the validations needed on the inputs
        If Asset_Name_TextBox.Text = String.Empty Then ' Validation checks on the asset name 
            MsgBox("Please fill in the asset's name")
        ElseIf IsNumeric(Asset_Name_TextBox.Text) Then
            MsgBox("numbers in names are not permitted")
        ElseIf Len(Asset_Name_TextBox.Text) < 4 Or Len(Asset_Name_TextBox.Text) > 15 Then
            MsgBox("Please enter a name that is between 4 and 15 characters")
        Else
            Create_New_Asset()
        End If

    End Sub

    Private Sub Create_New_Asset()
        ' this sub procedure saves the new asset created into the database 
        If DatabaseConnection() Then ' check the connection
            Dim SQLCMD As New OleDbCommand
            If ValueOfAssetID = -1 Then
                With SQLCMD
                    .Connection = Connection
                    .CommandText = "Insert into ASSET_TABLE (ASSET_NAME, ASSET_TYPE, ASSET_PRICE, ASSET_QUANTITY )" & "Values (@AssetName ,@AssetType ,@AssetPrice ,@AssetQuantity )" ' sql insertion of the asset details
                    .Parameters.AddWithValue("@AssetName", Asset_Name_TextBox.Text)
                    .Parameters.AddWithValue("@AssetType", Asset_Type)
                    .Parameters.AddWithValue("@AssetPrice", Asset_Price)
                    .Parameters.AddWithValue("@AssetQuantity", Asset_Quantity)
                    .ExecuteNonQuery()
                    .CommandText = "Select @@Identity"
                    ValueOfAssetID = .ExecuteScalar
                    Add_Asset_Auto_ID_Label.Text = ValueOfAssetID

                End With
                Connection.Close()
                Add_Asset_Auto_ID_Label.Text = "Automatically Generated"
                Display_Assets_Table()
            End If

        End If

    End Sub

    Private Sub Clear_Fields()
        ' this clears out all the text boxes 
        Asset_ID_Listbox.Items.Clear()
        Asset_Name_Listbox.Items.Clear()
        Asset_Type_Listbox.Items.Clear()
        Asset_Price_Listbox.Items.Clear()
        Search_Asset_ListBox.Items.Clear()
        Asset_Quantity_Listbox.Items.Clear()
        Assign_Asset_Name_TextBox.Clear()
        Asset_Name_TextBox.Clear()
        Search_Asset_Name_TextBox.Clear()
        Search_Asset_ID_TextBox.Clear()
        New_Asset_Name_TextBox.Clear()
        Asset_Price_NumericUpDown.Value = 0
        Asset_Quantity_NumericUpDown.Value = 0
        New_Asset_Price_NumericUpDown.Value = 0
        New_Asset_Quantity_NumericUpDown.Value = 0
        Asset_Type_ComboBox.SelectedIndex = -1
        New_Asset_Type_ComboBox.SelectedIndex = -1
        Assign_Asset_Local_ComboBox.SelectedIndex = -1
        Asset_Type_ComboBox.Items.Clear()
        New_Asset_Type_ComboBox.Items.Clear()
        Assign_Asset_Local_ComboBox.Items.Clear()
        ValueOfAssetID = -1
        Display_Asset_Name_Textbox.Clear()

    End Sub

    Private Sub Display_Assets_Table()
        ' this sub procedure displays the assets and its information
        If DatabaseConnection() Then
            Clear_Fields()
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select * From ASSET_TABLE " ' select all records from the database 

                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim AssetID, Name, Type, AssetPrice, Quantity As String '' these variables will store the data that is read by the recordset 
                    AssetID = DataReader("ASSET_ID")
                    Name = DataReader("ASSET_NAME")
                    Type = DataReader("ASSET_TYPE")
                    AssetPrice = DataReader("ASSET_PRICE")
                    Quantity = DataReader("ASSET_QUANTITY")

                    Asset_ID_Listbox.Items.Add(AssetID)
                    Asset_Name_Listbox.Items.Add(Name)
                    Asset_Type_Listbox.Items.Add(Type)
                    Asset_Price_Listbox.Items.Add(AssetPrice)
                    Asset_Quantity_Listbox.Items.Add(Quantity) ' displaying the data
                End While
                DataReader.Close()
            End With
            Connection.Close() ' closing the connection
        End If
        Load_Asset_Types()
        Load_Locals()

    End Sub
    Private Sub Load_Locals()
        ' this sub procedure loads all of the locals into the combo box so that it can be selected by the user 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select LOCAL_NAME From LOCAL_TABLE"

                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim Local As String ' this will hold the name read by the recordset
                    Local = DataReader("LOCAL_NAME")
                    If Not Assign_Asset_Local_ComboBox.Items.Contains(Local) Then
                        Assign_Asset_Local_ComboBox.Items.Add(Local) 'adding it into the combo box
                    End If
                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If
    End Sub
    Private Sub Asset_Management_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' this sub procedure is called when the form loads, it calls several procedures and disables some  buttons
        Display_Assets_Table()
        Delete_Asset_Button.Enabled = False
        Update_Asset_Button.Enabled = False
    End Sub
    Private Sub Load_Asset_Types()
        ' this sub procedure loads all of the different asset types into a combo box so it can be seleceted by the user
        Asset_Type_ComboBox.Items.Add("Piano")
        Asset_Type_ComboBox.Items.Add("Drum Kit")
        Asset_Type_ComboBox.Items.Add("Microphone")
        Asset_Type_ComboBox.Items.Add("Church Building")
        Asset_Type_ComboBox.Items.Add("Church Van")
        Asset_Type_ComboBox.Items.Add("Bass Guitar")
        Asset_Type_ComboBox.Items.Add("Lead Guitar")
        Asset_Type_ComboBox.Items.Add("Projector")

    End Sub

    Private Sub Asset_Type_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Asset_Type_ComboBox.SelectedIndexChanged
        ' this sub procedure stores the type of asset  the user has selected from the combo box 
        If Asset_Type_ComboBox.SelectedItem IsNot Nothing Then
            Asset_Type = Asset_Type_ComboBox.SelectedItem.ToString ' stores the asset type
        End If
    End Sub
    Private Sub Load_Asset_Name_Button_Click(sender As Object, e As EventArgs) Handles Load_Asset_Name_Button.Click
        'this sub procedure calls the sub procedure loads the assets name 
        If Asset_ID_Listbox.SelectedItem IsNot Nothing Then
            Load_Asset_Name()
        End If
    End Sub
    Private Sub Load_Asset_Name()
        ' this procedure loads the assets name into the textbox, if the asset's if its id is the same as the one selected by the user 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select ASSET_NAME " & "From ASSET_TABLE " & "Where ASSET_ID = @AssetID"

                .Parameters.AddWithValue("@AssetID", Selected_Asset_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim Name As String ' this store the name of the asset read by the record set 
                    Name = DataReader("ASSET_NAME")
                    Display_Asset_Name_Textbox.Text = Name 'displays the name
                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If
        Delete_Asset_Button.Enabled = True
    End Sub
    Private Sub Delete_Local_Asset()
        ' this sub procedure deletes all the linkages between the asset and local that corrospond to the asset being deleted
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Delete * " & "From LOCAL_ASSET_MAPPING_TABLE " & "Where ASSETS_ID = @DeleteID " ' deleting the record
                .Parameters.AddWithValue("@DeleteID", Selected_Asset_ID)
                .ExecuteNonQuery()
            End With
        End If
    End Sub
    Private Sub Delete_Asset()
        ' this is sub procedure deletes the asset from the database, where the asset id is the same as the id that was selected by the user  
        If DatabaseConnection() Then
            Delete_Local_Asset()
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Delete * " & "From ASSET_TABLE " & "Where ASSET_ID = @DeleteID " ' delete the asset
                .Parameters.AddWithValue("@DeleteID", Selected_Asset_ID)
                .ExecuteNonQuery()
            End With
            Connection.Close() ' close the connection
            Display_Assets_Table()
            Display_Asset_Name_Textbox.Clear()
            Delete_Asset_Button.Enabled = False
        End If
    End Sub

    Private Sub Delete_Asset_Button_Click(sender As Object, e As EventArgs) Handles Delete_Asset_Button.Click
        'this sub verifies if the user is sure they want to delete the asset
        Dim Delete As String ''creating the local variable delete
        Delete = MsgBox("Are you sure you would like to delete this Asset? Data will be permenantly deleted", vbExclamation + vbYesNo + vbDefaultButton2, "Delete Asset Confirmation")

        If Delete = vbYes Then ' if yes is selected
            Delete_Asset()

        ElseIf Delete = vbNo Then ' if no is selected 
            Display_Asset_Name_Textbox.Clear()
            Delete_Asset_Button.Enabled = False
        End If
    End Sub

    Private Sub Load_New_Asset_Types()
        ' this sub loads the different types of assets into a combo box
        New_Asset_Type_ComboBox.Items.Add("Piano")
        New_Asset_Type_ComboBox.Items.Add("Drum Kit")
        New_Asset_Type_ComboBox.Items.Add("Microphone")
        New_Asset_Type_ComboBox.Items.Add("Church Building")
        New_Asset_Type_ComboBox.Items.Add("Church Van")
        New_Asset_Type_ComboBox.Items.Add("Bass Guitar")
        New_Asset_Type_ComboBox.Items.Add("Lead Guitar")
        New_Asset_Type_ComboBox.Items.Add("Projector")
    End Sub
    Private Sub Load_Asset_Details_Button_Click(sender As Object, e As EventArgs) Handles Load_Asset_Details_Button.Click
        ' ' when this button is clicked it then calls these sub procedures 
        If Asset_ID_Listbox.SelectedItem IsNot Nothing Then
            New_Asset_Type_ComboBox.Items.Clear()

            Load_New_Asset_Types()
            Load_Assets_Details()
        End If

    End Sub
    Private Sub Load_Assets_Details()
        ' this sub procedure loads the current details about the asset into the text boxes so that they can be ammended to
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select * " & "From ASSET_TABLE " & "Where ASSET_ID = @AssetID"

                .Parameters.AddWithValue("@AssetID", Selected_Asset_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim AssetID, Name, Type, AssetPrice, Quantity As String ' these variables will hold the data read by the record set
                    AssetID = DataReader("ASSET_ID")
                    Name = DataReader("ASSET_NAME")
                    Type = DataReader("ASSET_TYPE")
                    AssetPrice = DataReader("ASSET_PRICE")
                    Quantity = DataReader("ASSET_QUANTITY")
                    New_Asset_Name_TextBox.Text = Name
                    New_Asset_Price_NumericUpDown.Value = Int(AssetPrice)
                    New_Asset_Quantity_NumericUpDown.Value = Int(Quantity) ' dsiplaying the data
                End While
                DataReader.Close()
            End With
            Connection.Close() ' close the connection
            Update_Asset_Button.Enabled = True
        End If
    End Sub

    Private Sub Update_Asset_Button_Click(sender As Object, e As EventArgs) Handles Update_Asset_Button.Click
        ' this sub carries out the validation checks on the users new input 
        If New_Asset_Name_TextBox.Text = String.Empty Then ' Validation checks on the asset name 
            MsgBox("Please fill in the asset's name")
        ElseIf IsNumeric(New_Asset_Name_TextBox.Text) Then
            MsgBox("numbers in names are not permitted")
        ElseIf Len(New_Asset_Name_TextBox.Text) < 4 Or Len(New_Asset_Name_TextBox.Text) > 15 Then
            MsgBox("Please enter a name that is between 4 and 15 characters")
        Else
            Update_Asset_Details()
        End If

    End Sub

    Private Sub Update_Asset_Details()
        ' this sub procedure updates the asset details that corrospong to the id that was selected by the user 
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Update ASSET_TABLE " & "Set ASSET_NAME = @AssetName, " & "ASSET_TYPE = @Type, " & "ASSET_PRICE = @Price, " & "ASSET_QUANTITY = @Quantity " & "Where ASSET_ID = @AssetID "

                .Parameters.AddWithValue("@AssetName", New_Asset_Name_TextBox.Text)
                .Parameters.AddWithValue("@Type", New_Asset_Type)
                .Parameters.AddWithValue(" @Price", New_Asset_Price)
                .Parameters.AddWithValue("@Quantity", New_Quantity)

                .Parameters.AddWithValue("@AssetID", Selected_Asset_ID)

                .ExecuteNonQuery()
                ' updataing the record 
                Update_Asset_Button.Enabled = False
            End With
            Connection.Close()
            Display_Assets_Table() ' display the details
        End If
    End Sub
    Private Sub Search_Asset_Name_Button_Click(sender As Object, e As EventArgs) Handles Search_Asset_Name_Button.Click
        ' this sub validates the input of the user and calls the sub that searches for the  asset
        Dim AssetName As String
        If Search_Asset_Name_TextBox.Text = String.Empty Then ' Validation checks on the asset name 
            MsgBox("Please fill in the asset's name")
        ElseIf IsNumeric(Search_Asset_Name_TextBox.Text) Then
            MsgBox("numbers in names are not permitted")
        Else
            AssetName = Search_Asset_Name_TextBox.Text
            Search_Asset_Name(AssetName)
        End If
    End Sub
    Private Sub Search_Asset_Name(AssetName As String)
        ' this sub procedure searches for an asset within the database 
        If DatabaseConnection() Then
            Clear_Fields()
            Get_Asset_Locals_Via_Name(AssetName)
            Dim SQLcmd As New OleDbCommand
            With SQLcmd
                .Connection = Connection
                .CommandText = "Select * " & "From ASSET_TABLE " & " Where ASSET_NAME Like @Asset "
                .Parameters.AddWithValue("@Asset", "%" & AssetName & "%") ' if a similar name is entered 
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim AssetID, Name, Type, AssetPrice, Quantity As String ' these will hold the values read by the record set 
                    AssetID = DataReader("ASSET_ID")
                    Name = DataReader("ASSET_NAME")
                    Type = DataReader("ASSET_TYPE")
                    AssetPrice = DataReader("ASSET_PRICE")
                    Quantity = DataReader("ASSET_QUANTITY")

                    If Not Asset_ID_Listbox.Items.Contains(AssetID) And Not Asset_Name_Listbox.Items.Contains(AssetName) And Not Asset_Type_Listbox.Items.Contains(Type) And Not Asset_Price_Listbox.Items.Contains(AssetPrice) And Not Asset_Quantity_Listbox.Items.Contains(Quantity) Then
                        Asset_ID_Listbox.Items.Add(AssetID)
                        Asset_Name_Listbox.Items.Add(Name)
                        Asset_Type_Listbox.Items.Add(Type)
                        Asset_Price_Listbox.Items.Add(AssetPrice)
                        Asset_Quantity_Listbox.Items.Add(Quantity) ' output the results
                    End If
                End While
                DataReader.Close()
            End With
            Connection.Close() ' connection close
        End If
    End Sub
    Private Sub Get_Asset_Locals_Via_Name(AssetName As String) ' passed in parameter 
        ' this sub procedure displays any locals that have been assinged to the asset 
        If DatabaseConnection() Then
            Clear_Fields()
            Dim SQLcmd As New OleDbCommand
            With SQLcmd
                .Connection = Connection
                .CommandText = "Select LOCAL_NAME,ASSET_NAME,LOCALS_ID " & "From LOCAL_TABLE,ASSET_TABLE,LOCAL_ASSET_MAPPING_TABLE" & " Where ASSET_NAME Like @Asset and LOCAL_ASSET_MAPPING_TABLE.LOCALS_ID = LOCAL_TABLE.LOCAL_ID  and LOCAL_ASSET_MAPPING_TABLE.ASSETS_ID = ASSET_TABLE.ASSET_ID" ' and LOCAL_ASSET_MAPPING_TABLE.ASSETS_ID = ASSET_TABLE.ASSET_ID " ' and LOCAL_TABLE.LOCAL_ID = MEMBERS_TABLE.LOCAL.ID " 'This is very similar to the StudentID Select query but this time we’re searching using the Surname field and the recordset returned can contain multiple records
                .Parameters.AddWithValue("@Asset", "%" & AssetName & "%") ' if a similar name is entere
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim Local As String ' holds the name of the local read by the record set 
                    Local = DataReader("LOCAL_NAME")
                    If Not Search_Asset_ListBox.Items.Contains(Local) Then
                        Search_Asset_ListBox.Items.Add(Local) ' adds the local into the listbox
                    End If
                End While
                DataReader.Close() ' closing the record set
            End With
        End If
    End Sub
    Private Sub Search_Asset_ID_Button_Click(sender As Object, e As EventArgs) Handles Search_Asset_ID_Button.Click
        'validates the user input into the search id textbox
        Dim ID As Integer
        If Search_Asset_ID_TextBox.Text = String.Empty Then ' Validation checks on the asset id
            MsgBox("Please fill in the asset's id")
        ElseIf Not IsNumeric(Search_Asset_ID_TextBox.Text) Then
            MsgBox("strings are not permitted in an ID")
        Else
            ID = Int(Search_Asset_ID_TextBox.Text)
            Search_Asset_ID(ID) 'passes the parameter through
        End If
    End Sub

    Private Sub Search_Asset_ID(Asset_ID As Integer) ' the passed in parameter
        'searches for record with matching id
        If DatabaseConnection() Then
            Clear_Fields()
            Get_Asset_Locals_Via_ID(Asset_ID)
            Dim SQLcmd As New OleDbCommand
            With SQLcmd
                .Connection = Connection
                .CommandText = "Select * " & "From ASSET_TABLE " & "Where ASSET_ID = @AssetID "


                .Parameters.AddWithValue("@AssetID", Asset_ID)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim AssetID, Name, Type, AssetPrice, Quantity As String ' these will sote the values read
                    AssetID = DataReader("ASSET_ID")
                    Name = DataReader("ASSET_NAME")
                    Type = DataReader("ASSET_TYPE")
                    AssetPrice = DataReader("ASSET_PRICE")
                    Quantity = DataReader("ASSET_QUANTITY")

                    Asset_ID_Listbox.Items.Add(AssetID)
                    Asset_Name_Listbox.Items.Add(Name)
                    Asset_Type_Listbox.Items.Add(Type)
                    Asset_Price_Listbox.Items.Add(AssetPrice)
                    Asset_Quantity_Listbox.Items.Add(Quantity) ' outputting the result
                End While
                DataReader.Close()
            End With
            Connection.Close() ' close the connection
        End If
    End Sub
    Private Sub Get_Asset_Locals_Via_ID(Asset_ID) ' passed in parameter 
        ' displays the local that uses the asset
        If DatabaseConnection() Then
            Clear_Fields() ' clears fields 
            Dim SQLcmd As New OleDbCommand
            With SQLcmd
                .Connection = Connection
                .CommandText = "Select LOCAL_NAME,ASSET_ID,LOCALS_ID " & "From LOCAL_TABLE,ASSET_TABLE,LOCAL_ASSET_MAPPING_TABLE" & " Where ASSET_ID = @AssetID and LOCAL_ASSET_MAPPING_TABLE.LOCALS_ID = LOCAL_TABLE.LOCAL_ID  and LOCAL_ASSET_MAPPING_TABLE.ASSETS_ID = ASSET_TABLE.ASSET_ID" ' and LOCAL_ASSET_MAPPING_TABLE.ASSETS_ID = ASSET_TABLE.ASSET_ID " ' and LOCAL_TABLE.LOCAL_ID = MEMBERS_TABLE.LOCAL.I
                .Parameters.AddWithValue("@AssetID", Asset_ID) '
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Dim Local As String ' store the name read by the record set 
                    Local = DataReader("LOCAL_NAME")
                    If Not Search_Asset_ListBox.Items.Contains(Local) Then
                        Search_Asset_ListBox.Items.Add(Local) ' local is displayed in the listbox
                    End If
                End While
                DataReader.Close()
            End With

        End If
    End Sub


    Private Sub Asset_Price_NumericUpDown_ValueChanged(sender As Object, e As EventArgs) Handles Asset_Price_NumericUpDown.ValueChanged
        ' stores the value of the asset price
        Asset_Price = Int(Asset_Price_NumericUpDown.Value.ToString)
    End Sub

    Private Sub Asset_Quantity_NumericUpDown_ValueChanged(sender As Object, e As EventArgs) Handles Asset_Quantity_NumericUpDown.ValueChanged
        ' stores the value of the asset quantity 
        Asset_Quantity = Int(Asset_Quantity_NumericUpDown.Value.ToString)
    End Sub

    Private Sub New_Asset_Type_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles New_Asset_Type_ComboBox.SelectedIndexChanged
        ' stores the value of the new asset type
        If New_Asset_Type_ComboBox.SelectedItem IsNot Nothing Then
            New_Asset_Type = New_Asset_Type_ComboBox.SelectedItem.ToString
        End If
    End Sub

    Private Sub New_Asset_Price_NumericUpDown_ValueChanged(sender As Object, e As EventArgs) Handles New_Asset_Price_NumericUpDown.ValueChanged
        ' stores the value of the new asset price
        New_Asset_Price = Int(New_Asset_Price_NumericUpDown.Value.ToString)
    End Sub

    Private Sub New_Asset_Quantity_NumericUpDown_ValueChanged(sender As Object, e As EventArgs) Handles New_Asset_Quantity_NumericUpDown.ValueChanged
        ' stores the value of the new asset quantity 
        New_Quantity = Int(New_Asset_Quantity_NumericUpDown.Value.ToString)
    End Sub

    Private Sub Search_Asset_ListBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Search_Asset_ListBox.SelectedIndexChanged
        'this sub procedure will store the name of the selected local from the list box that displays what locals use an asset 
        If Search_Asset_ListBox.SelectedItem IsNot Nothing Then
            Selected_Asset_Local = Search_Asset_ListBox.SelectedItem
            Fill_TextBoxes()
            Select_Local()
        End If
    End Sub
    Private Sub Select_Local()
        'this sub procedure stores the id of the local that was selected by the user from the listbox
        If DatabaseConnection() Then
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Select LOCAL_ID From LOCAL_TABLE " & "Where LOCAL_TABLE.LOCAL_NAME = @LocalName"
                .Parameters.AddWithValue("@LocalName", Selected_Asset_Local)
                Dim DataReader As OleDbDataReader = .ExecuteReader()
                While DataReader.Read
                    Asset_Local_ID = Int(DataReader("LOCAL_ID")) ' storing the local id 
                End While
                DataReader.Close()
            End With
            Connection.Close()
        End If


    End Sub

    Private Sub Assign_Asset_Local_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Assign_Asset_Local_ComboBox.SelectedIndexChanged
        ' this sub procedure stores the local id that corrosponds to the name of the local that was selected from the combo box
        Dim LocalName As String ' stores the name  selected from the combo box
        If Assign_Asset_Local_ComboBox.SelectedItem IsNot Nothing Then
            LocalName = Assign_Asset_Local_ComboBox.SelectedItem.ToString ' storing the local name
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
                Connection.Close() ' closing the connection
            End If
        End If
    End Sub

    Private Sub Assign_Asset_To_Local_Button_Click(sender As Object, e As EventArgs) Handles Assign_Asset_To_Local_Button.Click
        ' this sub procedure saves what local useses an asset to the asset local linking table 
        If Assign_Asset_Local_ComboBox.SelectedItem IsNot Nothing And Asset_ID_Listbox.SelectedItem IsNot Nothing Then
            If DatabaseConnection() Then
                Dim SQLCMD As New OleDbCommand
                If ValueOfLocal_AssetID = -1 Then
                    With SQLCMD
                        .Connection = Connection
                        .CommandText = "Insert into LOCAL_ASSET_MAPPING_TABLE (ASSETS_ID, LOCALS_ID )" & "Values (@AssetID ,@LocalID )" ' saving it into  the linking table 
                        .Parameters.AddWithValue("@AssetID", Selected_Asset_ID)
                        .Parameters.AddWithValue("@LocalID", Local_ID)
                        .ExecuteNonQuery()
                        .CommandText = "Select @@Identity"
                        ValueOfAssetID = .ExecuteScalar
                    End With
                    Connection.Close()
                    Display_Assets_Table() ' displays the data
                End If

            End If
        End If
    End Sub

    Private Sub Delete_Local_Asset_Button_Click(sender As Object, e As EventArgs) Handles Delete_Local_Asset_Button.Click
        ' this sub procedure deletes a local if they are no longer useing an asset
        If DatabaseConnection() Then
            Search_Asset_ListBox.Items.Remove(Selected_Asset_Local) ' removes it form the list box
            Dim SQLCMD As New OleDbCommand
            With SQLCMD
                .Connection = Connection
                .CommandText = "Delete * " & "From LOCAL_ASSET_MAPPING_TABLE " & "Where ASSETS_ID = @AssetID and LOCALS_ID = @LocalID "
                .Parameters.AddWithValue("@DeleteID", Selected_Asset_ID)
                .CommandText = "Delete * " & "From LOCAL_ASSET_MAPPING_TABLE " & "Where ASSETS_ID = @AssetID and LOCALS_ID = @LocalID "
                .Parameters.AddWithValue("@LocalID", Asset_Local_ID)
                ' deletes the link 
                .ExecuteNonQuery()
            End With
            Connection.Close() ' closes the connection
        End If
    End Sub
End Class