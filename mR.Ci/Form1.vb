' The Code you are viewing now is
' owned by SoftDevs Technologies, LTD.
' www.softdevs.puzl.com
' if you want to use them, Contact us @
' www.softdevs.puzl.com/contact-us
' Thank you for patronizing us

Imports Microsoft.VisualBasic.Devices
Imports System.Management
Imports System.IO
Imports mRCi.My.Resources
Public Class Form1
    Dim Dragging As Boolean
    Dim PointClicked As Point
    Dim FixedHardDriveName As String
    Dim GetFilename, Root, GetFilename2 As String, fold As DirectoryInfo, NumberOfRunningProcesses, SelectedNo As Int32
    Dim cpuUsage As New PerformanceCounter("Processor", "% Processor Time", "_Total")
    Dim request, strName As String
    Dim MyProcess As New Process()
    Dim isFormating As Boolean = False
    Dim Fle As Int16
    Dim s As String

    Private Sub Panel1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseDown

        If e.Button = MouseButtons.Left Then
            Dragging = True
            PointClicked = New Point(e.X, e.Y)
        Else
            Dragging = False
        End If
    End Sub

    Private Sub Panel1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseMove

        If Dragging Then
            Dim PointMoveTo As Point
            ' Find the current mouse position in screen coordinates.
            PointMoveTo = Me.PointToScreen(New Point(e.X, e.Y))

            'Compensate for the position the control was clicked
            PointMoveTo.Offset(-PointClicked.X, -PointClicked.Y)

            ' Move the Form.
            Me.Location = PointMoveTo
        End If
    End Sub

    Private Sub Panel1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseUp

        Dragging = False

    End Sub
    
    Enum InfoTypes
        OperatingSystemName
        ProcessorName
        AmountOfMemory
        VideocardName
    End Enum

    Public Function GetInfo(ByVal InfoType As InfoTypes) As String
        Dim Info As New ComputerInfo : Dim Value, vganame, proc As String
        Dim searcher As New Management.ManagementObjectSearcher( _
            "root\CIMV2", "SELECT * FROM Win32_VideoController")
        Dim searcher1 As New Management.ManagementObjectSearcher( _
            "SELECT * FROM Win32_Processor")
        If InfoType = InfoTypes.OperatingSystemName Then
            Value = Info.OSFullName
        ElseIf InfoType = InfoTypes.ProcessorName Then
            For Each queryObject As ManagementObject In searcher1.Get
                proc = queryObject.GetPropertyValue("Name").ToString
            Next
            Value = proc
        ElseIf InfoType = InfoTypes.AmountOfMemory Then
            Value = Math.Round((((CDbl(Convert.ToDouble(Val(Info.TotalPhysicalMemory))) / 1024)) / 1024), 2) & " MB"
        ElseIf InfoType = InfoTypes.VideocardName Then
            For Each queryObject As ManagementObject In searcher.Get
                vganame = queryObject.GetPropertyValue("Name").ToString
            Next
            Value = vganame
        End If
        Return Value
    End Function
    
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '------------Check if Software is Registered or Not----------------
        If My.Settings.Registered = False Then
            My.Settings.Loads -= 1
            lblLoads.Visible = True
            lblLoads.Text = My.Settings.Loads.ToString + " more loads remaining."
            lblRegistered.Text = "Registered : False"
            btnRegister.Enabled = True
            My.Settings.Save()
            If My.Settings.Loads <= -1 Then
                If MessageBox.Show("Your 10x Access of mR.Ci Application has been Expired. Please register to continue using this software.", "Trial Expired", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                    Register.Show()
                    Me.Close()
                Else
                    End
                End If
            End If
        ElseIf My.Settings.Registered = True Then
            lblLoads.Visible = False
            lblRegistered.Text = "Registered : True"
            btnRegister.Enabled = False
            My.Settings.Save()
        End If
        '------------Cpu Usage Timer---------------------------------------
        tmrProcess.Enabled = True
        '------------Show All Drives for Disk Space Feature----------------
        Dim dirs() As String = System.IO.Directory.GetLogicalDrives
        Dim i As Integer
        For i = dirs.GetLowerBound(0) To dirs.GetUpperBound(0)
            cbDrives.Items.Add(dirs(i))
        Next
        '------------Processes Running-------------------------------------
        For Each Drives In My.Computer.FileSystem.Drives
            ComboBox3.Items.Add(Drives)
        Next
        ComboBox3.SelectedIndex = 0
        ComboBox3.Text = cbDrives.SelectedItem
        ' --------------Drive Formatter------------------------------------
        If My.User.IsInRole("administrators") = True Then
            For Each AttachedDrive In Environment.GetLogicalDrives()
                Dim CompDrive As New DriveInfo(AttachedDrive)
                If Not CompDrive.Name = "A:\" And Not CompDrive.Name = "B:\" Then
                    If (CompDrive.DriveType = DriveType.Removable Or CompDrive.DriveType = DriveType.Fixed) Then
                        comboboxDrives.Items.Add(CompDrive.Name)
                    End If
                End If
            Next
            isFormating = False
            comboboxDrives.SelectedIndex = 0
            Dim computerdrive As New DriveInfo(comboboxDrives.SelectedItem.ToString())
            TextboxVolumeLabel.Text = computerdrive.VolumeLabel
        Else
            MessageBox.Show("Administrators prevelegies required.")
            Me.Close()
        End If
    End Sub

    Private Sub ButtonCheck_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCheck.Click
        TextBox3.Text = "True"
        On Error Resume Next
        'clear the list before adding other process
        ListView1.Items.Clear()
        For Each Proc In Process.GetProcesses
            Application.DoEvents()
            'get the root directory for a file
            Root = Path.GetPathRoot(Proc.MainModule.FileName)
            'check if process root directory is the same as the one you want to get files that are running on it
            If Root = ComboBox3.Text Then
                Application.DoEvents()
                NumberOfRunningProcesses += 1
                'add the fullpath of the running process and the filename
                ListView1.Items.Add(Proc.MainModule.FileName).SubItems.Add(Proc.ProcessName).ForeColor = Color.Pink

            End If
        Next
        Root = ComboBox3.Text
        For Each FileToCheck In Directory.GetFiles(Root, "*.*", SearchOption.AllDirectories)
            Application.DoEvents()
            FileInUse(FileToCheck)
            Application.DoEvents()
        Next
        Label5.Text = "Process Running : " & NumberOfRunningProcesses.ToString
        NumberOfRunningProcesses = 0
    End Sub

    Function FileInUse(ByVal sPath As String) As Boolean
        Try
            Fle = CShort(FreeFile())
            'check if file is in use
            FileOpen(Fle, sPath, OpenMode.Binary, OpenAccess.ReadWrite, OpenShare.LockReadWrite)
            FileClose(Fle)
        Catch
            ListView1.Items.Add(sPath).SubItems.Add("is in use.").ForeColor = Color.Blue
        End Try
    End Function

    Private Sub ButtonEndProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonEndProcess.Click
        'check if there are checked item(s)
        If Not ListView1.CheckedItems.Count = Nothing Then
            If ListView1.CheckedItems.Count >= 1 Then
            End If
            'get all checked process
            For Each selectedprocess In ListView1.CheckedItems
                'get the filename of the checked item
                GetFilename = selectedprocess.ToString.Remove(0, selectedprocess.ToString.IndexOf("{") + 1).Replace("}", "").ToString
                GetFilename2 = Path.GetFileNameWithoutExtension(GetFilename)
                For Each pro In Process.GetProcessesByName(GetFilename2)
                    'end all processes that are selected
                    'MsgBox(GetFilename.ToString)
                    pro.Kill()
                    ListView1.Items.RemoveAt(ListView1.CheckedIndices(0))
                Next
            Next
            MsgBox("Process was ended successfully")
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonReload.Click
        TextBox3.Text = False
        ListView1.Items.Clear()
        ComboBox3.Items.Clear()
        For Each Drives In My.Computer.FileSystem.Drives
            ComboBox3.Items.Add(Drives)
        Next
    End Sub

    Private Sub CbDrives_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbDrives.SelectedIndexChanged
        Try
            tbFreeSpace.Text = Math.Round((((CDbl(Convert.ToDouble(Val(My.Computer.FileSystem.Drives(cbDrives.SelectedIndex).AvailableFreeSpace))) / 1024)) / 1024) / 1024, 2) & " GB"
            tbTotalSpace.Text = Math.Round((((CDbl(Convert.ToDouble(Val(My.Computer.FileSystem.Drives(cbDrives.SelectedIndex).TotalSize))) / 1024)) / 1024) / 1024, 2) & " GB"
        Catch ex As Exception
            MsgBox("Please Select any Drive")
        End Try
    End Sub

    Private Sub tmrProcess_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrProcess.Tick
        lblProcess.Text = cpuUsage.NextValue.ToString + "%"
        pbRam.Value = pcRam.NextValue
        lblRam.Text = pbRam.Value.ToString + "%"
    End Sub

    Private Sub btnScan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnScan.Click
        TextBox2.Text = "True"
        WindowsName.Text = (GetInfo(InfoTypes.OperatingSystemName))
        ProcessorName.Text = (GetInfo(InfoTypes.ProcessorName))
        TotalRAM.Text = (GetInfo(InfoTypes.AmountOfMemory))
        VideocardName.Text = (GetInfo(InfoTypes.VideocardName))
        ComputerName.Text = My.Computer.Name.ToString
        WindowsVersion.Text = System.Environment.OSVersion.ToString
        WindowsPlatform.Text = My.Computer.Info.OSPlatform.ToString
        WindowsDirectory.Text = System.Environment.SystemDirectory.ToString
        InstalledLanguage.Text = My.Computer.Info.InstalledUICulture.ToString
        User.Text = System.Environment.UserName.ToString
        Capslock.Text = My.Computer.Keyboard.CapsLock
        Numlock.Text = My.Computer.Keyboard.NumLock
        Mouse.Text = My.Computer.Mouse.WheelExists
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        TextBox2.Text = "False"
        ComputerName.Clear()
        WindowsName.Clear()
        WindowsVersion.Clear()
        WindowsPlatform.Clear()
        WindowsDirectory.Clear()
        InstalledLanguage.Clear()
        ProcessorName.Clear()
        TotalRAM.Clear()
        VideocardName.Clear()
        User.Clear()
        Capslock.Clear()
        Numlock.Clear()
        Mouse.Clear()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If TextBox2.Text = "True" Then
            Me.SaveFileDialog1.FileName = "My_Computer_Info"
            Me.SaveFileDialog1.Filter = "Text Files (*.txt)|*.txt"
            If Me.SaveFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                If Not Me.SaveFileDialog1.FileName = Nothing Then
                    Dim myWriter As New IO.StreamWriter(Me.SaveFileDialog1.FileName)
                    myWriter.WriteLine("This is Your Custom Computer Information that mR.Ci Scanned")
                    myWriter.WriteLine(" ")
                    myWriter.WriteLine("Computer Name : " & ComputerName.Text)
                    myWriter.WriteLine("Windows Name : " & WindowsName.Text)
                    myWriter.WriteLine("Windows Version : " & WindowsVersion.Text)
                    myWriter.WriteLine("Windows Platform : " & WindowsPlatform.Text)
                    myWriter.WriteLine("Windows Directory : " & WindowsDirectory.Text)
                    myWriter.WriteLine("Installed Language : " & InstalledLanguage.Text)
                    myWriter.WriteLine("Processor Name : " & ProcessorName.Text)
                    myWriter.WriteLine("Total Memory (RAM) : " & TotalRAM.Text)
                    myWriter.WriteLine("Videocard Name : " & VideocardName.Text)
                    myWriter.WriteLine(" ")
                    myWriter.WriteLine("(c) SoftDevs Technologies 2013")
                    myWriter.Close()
                    MsgBox("Save Data to " & Me.SaveFileDialog1.FileName, MsgBoxStyle.Information, "Saved Successfully")
                End If
            End If
        ElseIf TextBox2.Text = "False" Then
            MsgBox("mR.Ci can't process the save option. There is no any data that has been scanned", MsgBoxStyle.Exclamation, "Error 404 : Data not found.")
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If LabelProgress.Text.Contains("Volume Serial") = False Then

            'check if format was aborted
            If LabelProgress.Text.Contains("Format aborted.") = False Then
                LabelInfo.Text = "Please wait. The formater is currently formatting the selected drive..."
                Application.DoEvents()
                'get the current progress
                LabelProgress.Text = MyProcess.StandardOutput.ReadLine().ToString()
                'disable the format button
                ButtonFormat.Enabled = False
                isFormating = True

            Else
                Timer1.Enabled = False
                LabelInfo.Text = "Format aborted."
                MessageBox.Show(LabelInfo.Text, "Error")
                'clear labelprogress
                LabelProgress.Text = String.Empty
                isFormating = False
                ButtonFormat.Enabled = True
            End If

            If LabelProgress.Text.Contains("Enter current volume label") = True Then
                Dim stremwriter As StreamWriter
                stremwriter = MyProcess.StandardInput
                stremwriter.Write(FixedHardDriveName)
                stremwriter.Close()
            End If

        Else
            Timer1.Enabled = False
            LabelInfo.Text = "Done Formating"
            LabelProgress.Text = String.Empty
            MessageBox.Show(LabelInfo.Text)
            isFormating = False
            ButtonFormat.Enabled = True
        End If

    End Sub

    Private Sub buttonFormat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonFormat.Click
        Dim SelectedDrive As New DriveInfo(comboBoxDrives.Text)
        If MessageBox.Show("You really want to format this drive : " & comboboxDrives.SelectedItem & " " & TextboxVolumeLabel.Text & "?", "Info", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            If SelectedDrive.DriveType = DriveType.Removable Then
                'check if quick format checkbox is checked
                strName = TextboxVolumeLabel.Text.Trim()
                'note this line can also work on LIMITED/STANDARD USER
                request = "Format /v:" + strName + " /x /Q /backup " + SelectedDrive.Name.Replace("\", " ")
                MyProcess.StartInfo.RedirectStandardOutput = True
                MyProcess.StartInfo.RedirectStandardError = True
                MyProcess.StartInfo.CreateNoWindow = True
                MyProcess.StartInfo.UseShellExecute = False
                'this try statement is here to avoid the application to crash,
                'when the is no enough space, read-only location
                Try
                    File.WriteAllText("batman.bat", request + Environment.NewLine)
                Catch m As IOException
                    MessageBox.Show(m.Message.ToString())
                    ButtonFormat.Enabled = True
                    Return
                End Try
                MyProcess.StartInfo.FileName = "batman.bat"
                'START the process
                MyProcess.Start()
                Timer1.Enabled = True
            End If

        End If
        If SelectedDrive.DriveType = DriveType.Fixed Then
            MyProcess.StartInfo.RedirectStandardOutput = True
            MyProcess.StartInfo.RedirectStandardError = True
            MyProcess.StartInfo.CreateNoWindow = True
            MyProcess.StartInfo.UseShellExecute = False
            Try
                File.WriteAllText("batman.bat", "Format /q /x /v:" & textBoxVolumeLabel.Text & " " & comboBoxDrives.SelectedItem.ToString())
            Catch m As IOException
                MessageBox.Show(m.Message.ToString())
                buttonFormat.Enabled = True
                Return
            End Try
            MyProcess.StartInfo.FileName = "batman.bat"
            'START the process
            MyProcess.Start()
            Timer1.Enabled = True
        End If
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        If isFormating = True Then
            e.Cancel = False
        Else
            If (File.Exists("batman.bat") = True) Then
                'delete the batman.bat file
                File.Delete("batman.bat")
            End If
        End If
        'check if the application is closed by a taskmanager
        If e.CloseReason = CloseReason.TaskManagerClosing Then
            If File.Exists("batman.bat") = True Then
                'delete the batman.bat file
                File.Delete("batman.bat") '
            End If
        End If
    End Sub

    Private Sub comboBoxDrives_SelectionChangeCommitted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles comboboxDrives.SelectionChangeCommitted
        Dim SelectedDrive As New DriveInfo(comboBoxDrives.Text)
        If SelectedDrive.IsReady = True Then
            If SelectedDrive.VolumeLabel = String.Empty Then
                textBoxVolumeLabel.Text = "Removable Disc"
                FixedHardDriveName = SelectedDrive.VolumeLabel
            Else
                FixedHardDriveName = SelectedDrive.VolumeLabel
                textBoxVolumeLabel.Text = SelectedDrive.VolumeLabel
            End If

        End If

    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        End
    End Sub

    Private Sub btnRegister_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRegister.Click
        Register.Show()
    End Sub

    Private Sub PictureBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox3.Click
        My.Settings.Loads = 10
        My.Settings.Registered = False
        MsgBox("Admin Success")
        My.Settings.Save()
    End Sub
End Class
