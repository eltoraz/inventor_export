' <IsStraightVb>True</IsStraightVb>
Imports System.Diagnostics

'call the DMT to add/update the specified part/revision (etc.)
Public Class DMT
    Public Shared dmt_loc As String = "C:\Epicor\ERP10.1Client\Client\DMT.exe"
    Public Shared dmt_working_path As String = "I:\Cadd\_iLogic\Export\"

    'TODO: change in production to DMT user/password/environment
    Private Shared username As String = "DMT_USERNAME"
    Private Shared password As String = "DMT_PASSWORD"
    Private Shared configfile As String = "EpicorPilot10"
    Private Shared connection As String = "net.tcp://CHERRY/EpicorPilot10"
    Private Shared dmt_base_args As String = "-NoUI -User=" & username & " -Pass=" & password & " -ConnectionURL=""" & connection & """ -ConfigValue=""" & configfile & """"

    'Run the DMT to import the specified CSV into Epicor
    Public Shared Sub dmt_import(csv As String, filename As String)
        Dim psi As New ProcessStartInfo(dmt_loc)
        psi.RedirectStandardOutput = True
        psi.WindowStyle = ProcessWindowStyle.Hidden
        psi.UseShellExecute = False

        psi.Arguments = dmt_base_args & " -Import=""" & csv & """ -Source="""
        psi.Arguments = psi.Arguments & filename & """ -Add=True -Update=True"

        Dim msg_succ As String = csv & " successfully imported into Epicor!"

        exec_dmt(psi, msg_succ)
    End Sub

    'use the DMT to export data from Epicor based on existing BAQs
    'the results of the queries is stored in the paired CSV files for later reading
    Public Shared Sub dmt_export()
        Dim export_path = dmt_working_path & "ref\"

        'Mapping of queries in Epicor and the corresponding output files
        Dim query_map As New Dictionary(Of String, String)
        query_map.Add("DMTVendorExport", "Vendors.csv")
        query_map.Add("DMTProdCode", "ProdCode.csv")
        query_map.Add("DMTClasSID", "ClassID.csv")

        Dim psi As New ProcessStartInfo(dmt_loc)
        psi.RedirectStandardOutput = True
        psi.WindowStyle = ProcessWindowStyle.Hidden
        psi.UseShellExecute = False

        For Each kvp As KeyValuePair(Of String, String) in query_map
            psi.Arguments = dmt_base_args & " -Export -BAQ=""" & kvp.Key
            psi.Arguments = psi.Arguments & """ -Target=""" & export_path & kvp.Value & """"

            msg_succ = "Successfully exported " & kvp.Key & " from Epicor"
            exec_dmt(psi, msg_succ)
        Next
    End Sub

    Public Shared Sub exec_dmt(psi As ProcessStartInfo, msg_succ As String)
        Dim dmt As Process
        dmt = Process.Start(psi)
        'Wait 30s (worst case) for DMT to exit - if it takes this long, something's wrong
        dmt.WaitForExit(30000)

        Dim resultmsg As String
        If Not dmt.HasExited Then
            resultmsg = "Warning: DMT has not finished after 30 seconds"
        ElseIf dmt.ExitCode = 0 Then
            resultmsg = msg_succ
        Else
            resultmsg = "Error: DMT exited with code " & dmt.ExitCode
        End If

        Dim event_time = DateTime.Now
        resultmsg = event_time.ToString("HH:mm:ss") & ": " & resultmsg

        dmt_log_event(resultmsg)
    End Sub

    Public Shared Function write_csv(csv_name As String, fields As String, data As String)
        Dim fso, file_name, csv

        'Open the CSV file (note: this will overwrite the file if it exists!)
        fso = CreateObject("Scripting.FileSystemObject")
        file_name = dmt_working_path & csv_name
        csv = fso.OpenTextFile(file_name, 2, True, -2)

        'Write field headers & data to file
        csv.WriteLine(fields)
        csv.WriteLine(data)
        csv.Close()

        'need to return the full path & filename to pass to DMT
        Return file_name
    End Function

    Public Shared Sub dmt_log_event(msg As String)
        Dim fso, file_name, log_file
        Dim log_date = DateTime.Now

        fso = CreateObject("Scripting.FileSystemObject")
        file_name = dmt_working_path & log_date.ToString("yyyyMMdd") & "_dmtlog.txt"
        log_file = fso.OpenTextFile(file_name, 8, True, -2)

        log_file.WriteLine(msg)
        log_file.Close()
    End Sub
End Class
