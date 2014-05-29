Public Class Register

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = My.Settings.Key1 Or TextBox1.Text = My.Settings.Key2 Or TextBox1.Text = My.Settings.Key3 Or TextBox1.Text = My.Settings.Key4 Or TextBox1.Text = My.Settings.Key5 Then
            My.Settings.Registered = True
            MsgBox("mR.Ci Application has been successfully registered. Thank you for using our product.", MsgBoxStyle.Information, "Registration Successfull")
            Form1.btnRegister.Enabled = False
            Form1.lblRegistered.Text = "Registered : True"
            Form1.lblLoads.Visible = False
            My.Settings.Save()
            Form1.Show()
            Me.Close()
        Else
            Form1.lblRegistered.Text = "Registered : False"
            Form1.btnRegister.Enabled = True
            Form1.lblLoads.Visible = True
            My.Settings.Registered = False
            My.Settings.Save()
            MsgBox("License is not valid or not registered in our system. Please Try Again", MsgBoxStyle.Exclamation, "Registration Failed")
        End If
    End Sub
End Class