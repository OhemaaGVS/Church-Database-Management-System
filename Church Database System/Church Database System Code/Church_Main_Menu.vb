Public Class Church_Main_Menu
    Private Sub Member_Management_Button_Click(sender As Object, e As EventArgs) Handles Member_Management_Button.Click
        ' when this button is clicked it opens the member management form 
        Member_Management_Form.ShowDialog()
    End Sub

    Private Sub Local_Management_Button_Click(sender As Object, e As EventArgs) Handles Local_Management_Button.Click
        ' when this button is clicked it opens the local management form 
        Local_Management_Form.ShowDialog()
    End Sub
    Private Sub Area_Management_Button_Click(sender As Object, e As EventArgs) Handles Area_Management_Button.Click
        ' when this button is clicked it opens the area management form 
        Area_Management_Form.ShowDialog()
    End Sub
    Private Sub District_Management_Button_Click(sender As Object, e As EventArgs) Handles District_Management_Button.Click
        ' when this button is clicked it opens the district management form 
        District_Management_Form.ShowDialog()
    End Sub
    Private Sub Service_Management_Button_Click(sender As Object, e As EventArgs) Handles Service_Management_Button.Click
        ' when this button is clicked it opens the service management form 
        Service_Management_Form.ShowDialog()
    End Sub
    Private Sub Asset_Management_Button_Click(sender As Object, e As EventArgs) Handles Asset_Management_Button.Click
        ' when this button is clicked it opens the asset management form 
        Asset_Management_Form.ShowDialog()
    End Sub
End Class
