Imports System
Imports System.Management
Imports System.IO

Public Class Form1

    Private BLOCKEDPATH As String = ""
    Dim processEventStartWatcher As ManagementEventWatcher

    Private Sub Form1_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        RemoveHandler processEventStartWatcher.EventArrived, AddressOf EventArrived
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As EventArgs) Handles Button1.Click
        Dim fbd As New FolderBrowserDialog

        If fbd.ShowDialog = DialogResult.OK Then
            TextBox1.Text = fbd.SelectedPath
        End If
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text <> "" AndAlso New DirectoryInfo(TextBox1.Text).Exists Then
            BLOCKEDPATH = TextBox1.Text

            Dim procWatcher As New Threading.Thread(AddressOf CreationEventWatcherPolling)
            procWatcher.Start()

            Button2.Enabled = False
            Button4.Enabled = True
        End If
    End Sub
    
    Private Sub Button4_Click(sender As System.Object, e As EventArgs) Handles Button4.Click
        ForceStopWatcher()
    End Sub

    Private Delegate Sub ForceStopWatcherDelegate()
    Private Sub ForceStopWatcher()
        BeginInvoke(New ForceStopWatcherDelegate(AddressOf ForceStopWatcherSub))
    End Sub
    Private Sub ForceStopWatcherSub()
        If processEventStartWatcher IsNot Nothing Then
            processEventStartWatcher.Stop()

            Button2.Enabled = True
            Button4.Enabled = False
        End If
    End Sub

    Private Sub CreationEventWatcherPolling()
        Dim query As New WqlEventQuery("__InstanceCreationEvent", New TimeSpan(1000000), "TargetInstance isa ""Win32_Process""")

        processEventStartWatcher = New ManagementEventWatcher(query)
        AddHandler processEventStartWatcher.EventArrived, AddressOf EventArrived
        processEventStartWatcher.Start()
    End Sub
    
    Private Sub EventArrived(ByVal sender As Object, ByVal e As EventArrivedEventArgs)
        Try
            Dim targetInstance As ManagementBaseObject = e.NewEvent("TargetInstance")

            If targetInstance("Name").ToString.ToLower.Contains(".exe") AndAlso targetInstance("ExecutablePath").ToString.ToLower.Contains(BLOCKEDPATH.ToLower) Then
                Process.GetProcessById(CInt(targetInstance("Handle"))).Kill()
                MessageBox.Show("Blocked!")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        

    End Sub

End Class
