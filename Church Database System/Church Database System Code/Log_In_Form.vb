Public Class Log_In_Form
    Private ActualUserName As String = "TheCopManagement" ' this varaible stores the value of the actual username that the user will need to type in
    Private ActualPassword As String = "TheCop2021" ' this varaible stores the value of the actual password that the user will need to type in
    Private UserName As String ' this is the variable will store the username that the user types in
    Private Password As String ' this is the variable will store the password that the user types in
    Private Sub Clear_Button_Click(sender As Object, e As EventArgs) Handles Clear_Button.Click
        'when this button ia clicked it will clear the textboxes
        Username_TextBox.Clear()
        Password_TextBox.Clear()
    End Sub

    Private Sub Log_In_Button_Click(sender As Object, e As EventArgs) Handles Log_In_Button.Click
        'when this button is clicked it will check if a user name and password has been inputted. if there has then it will clal the login sub procedure
        If Username_TextBox.Text = String.Empty Then ' checks if textbox has some text inside 
            MsgBox("Please enter a user name")
        ElseIf Password_TextBox.Text = String.Empty Then ' checks if textbox has some text inside 
            MsgBox("Please enter a password")
        Else
            UserName = Username_TextBox.Text
            Password = Password_TextBox.Text ' assigning the values that the user typed in to the variables 


            Login() ' calling the login sub procedure
          
        End If
    End Sub
    Private Sub Login()
        '  when this sub procedure is called it will check if the password and username match the actual password and username. if it is then it will open the main menu 
        If UserName = ActualUserName And Password = ActualPassword Then ' if they are equal
            MsgBox("Welcome User")
            Password_TextBox.Clear()
            Username_TextBox.Clear()
            Church_Main_Menu.ShowDialog() ' open the main menu 
        ElseIf UserName <> ActualUserName Or Password <> ActualPassword Then ' if eiether the password or username is incorrect
            MsgBox("Invalid Log In") ' tells the user its incorrect 
        End If
    End Sub


    Private Sub Log_In_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class