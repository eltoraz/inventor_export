'call the DMT to add/update the specified part/revision (etc.)
Public Class DMT
    Public Shared Function exec_DMT(csv As String, filename As String)
        'Call the DMT on the passed CSV file
        Dim dmt_loc = "C:\Epicor\ERP10.1Client\Client\DMT.exe"
        Dim psi As New System.Diagnostics.ProcessStartInfo(dmt_loc)
        psi.RedirectStandardOutput = True
        psi.WindowStyle = ProcessWindowStyle.Hidden
        psi.UseShellExecute = False

        'TODO: change in production to DMT user/password/environment
        Dim username, password, configfile, connection As String
        username = "DMT_USERNAME"
        password = "DMT_PASSWORD"
        configfile = "EpicorPilot10"
        connection = "net.tcp://CHERRY/EpicorPilot10"

        psi.Arguments = "-NoUI=True -Import=""" & csv & """ -Source=""" & filename
        psi.Arguments = psi.Arguments & """ -Add=True -Update=True -user=" & username
        psi.Arguments = psi.Arguments & " -pass=" & password & " -ConnectionUrl="""
        psi.Arguments = psi.Arguments & connection & """ -ConfigValue="""
        psi.Arguments = psi.Arguments & configfile & """"

        Dim dmt As System.Diagnostics.Process
        dmt = System.Diagnostics.Process.Start(psi)
        dmt.WaitForExit()

        Dim msgSuccess As String = csv & " successfully imported into Epicor!"
        Dim msgFailure As String = "Error importing part into Epicor!"

        Dim resultmsg As String
        If dmt.ExitCode = 0 Then
            resultmsg = msgSuccess
        Else
            resultmsg = msgFailure
        End If

        Return resultmsg
    End Function
End Class
