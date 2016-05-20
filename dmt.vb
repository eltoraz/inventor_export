'call the DMT to add/update the specified part/revision (etc.)
Public Class DMT
    Public Shared csv_path As String = "I:\Cadd\_iLogic\Export\"
    Private dmt_log_path As String = "I:\Cadd\_iLogic\Export\"

    Public Shared Sub exec_DMT(csv As String, filename As String)
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

        Dim event_time = DateTime.Now
        resultmsg = event_time.ToString("HHmmss") & ": " & resultmsg

        dmt_log_event(resultmsg)
    End Sub

    Public Shared Function write_csv(csv_name As String, fields As String, data As String)
        Dim fso, file_name, csv

        'Open the CSV file (note: this will overwrite the file if it exists!)
        fso = CreateObject("Scripting.FileSystemObject")
        file_name = csv_path & csv_name
        csv = fso.OpenTextFile(file_name, 2, True, -2)

        'Write field headers & data to file
        csv.WriteLine(fields)
        csv.WriteLine(data)
        csv.Close()

        'need to return the full path & filename to pass to DMT
        Return file_name
    End Function

    Private Sub dmt_log_event(msg As String)
        Dim fso, file_name, log_file
        Dim log_date = DateTime.Now

        fso = CreateObject("Scripting.FileSystemObject")
        file_name = dmt_log_path & log_date.ToString("yyyyMMdd") & "_dmtlog.txt"
        log_file = fso.OpenTextFile(file_name, 8, True, -2)

        log_file.WriteLine(msg)
        log_file.Close()
    End Sub
End Class
