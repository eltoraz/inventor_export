' <IsStraightVb>True</IsStraightVb>
Imports System.Text.RegularExpressions
Imports System.Diagnostics
Imports System.IO

'call the DMT to add/update the specified part/revision (etc.)
Public Class DMT
    Public Shared dmt_loc As String = "C:\Epicor\ERP10.1Client\Client\DMT.exe"
    Public Shared dmt_working_path As String = "I:\Cadd\_Epicor\"

    'TODO: change in production to DMT user/password/environment
    Private username As String = "DMT_USERNAME"
    Private password As String = "DMT_PASSWORD"
    Private configfile As String = "EpicorPilot10"
    Private connection As String = "net.tcp://CHERRY/EpicorPilot10"
    Private dmt_base_args As String = "-NoUI -User=" & username & " -Pass=" & password & " -ConnectionURL=""" & connection & """ -ConfigValue=""" & configfile & """"

    'Run the DMT to import the specified CSV into Epicor
    'pass along the return code from the DMT (-1 if it timed out)
    'if the 3rd arg is true, have DMT update the part in Epicor
    '(just try adding it as a new entry otherwise)
    Public Function dmt_import(csv As String, filename As String, update_on As Boolean) _
                               As Integer
        Dim psi As New ProcessStartInfo(dmt_loc)
        psi.RedirectStandardOutput = True
        psi.WindowStyle = ProcessWindowStyle.Hidden
        psi.UseShellExecute = False

        psi.Arguments = dmt_base_args & " -Import=""" & csv & """ -Source="""
        psi.Arguments = psi.Arguments & filename & """ -Add"

        If update_on Then
            psi.Arguments = psi.Arguments & """ -Update"
        End If

        Dim msg_succ As String = "Successfully imported into Epicor!"

        Return exec_dmt(psi, csv, msg_succ)
    End Function

    'use the DMT to export data from Epicor based on existing BAQs
    'the results of the queries is stored in the paired CSV files for later reading
    'pass along the return code from the DMT (-1 if it timed out)
    Public Sub dmt_export()
        Dim export_path = dmt_working_path & "ref\"

        'Mapping of queries in Epicor and the corresponding output files
        Dim query_map As New Dictionary(Of String, String)
        query_map.Add("DMTProdCode", "ProdCode.csv")
        query_map.Add("DMTClasSID", "ClassID.csv")

        Dim psi As New ProcessStartInfo(dmt_loc)
        psi.RedirectStandardOutput = True
        psi.WindowStyle = ProcessWindowStyle.Hidden
        psi.UseShellExecute = False

        For Each kvp As KeyValuePair(Of String, String) in query_map
            psi.Arguments = dmt_base_args & " -Export -BAQ=""" & kvp.Key
            psi.Arguments = psi.Arguments & """ -Target=""" & export_path & kvp.Value & """"

            msg_succ = "Successfully exported CSV from Epicor"
            exec_dmt(psi, kvp.Key, msg_succ)
        Next
    End Sub

    'return -1 if DMT times out, otherwise pass on DMT's return value (as per
    'convention, 0 is success and >0 is an error, though I've only ever seen 1)
    Public Function exec_dmt(psi As ProcessStartInfo, prefix As String, msg_succ As String) _
                             As Integer
        Dim dmt As Process
        dmt = Process.Start(psi)
        'Wait 30s (worst case) for DMT to exit - if it takes this long, something's wrong
        dmt.WaitForExit(30000)

        Dim resultmsg As String
        Dim ret_value As Integer
        If Not dmt.HasExited Then
            resultmsg = "Warning: DMT has not finished after 30 seconds"
            ret_value = -1
        ElseIf dmt.ExitCode = 0 Then
            resultmsg = msg_succ
            ret_value = dmt.ExitCode
        Else
            resultmsg = "Error: DMT exited with code " & dmt.ExitCode
            ret_value = dmt.ExitCode
        End If

        dmt_log_event(prefix, resultmsg)
        Return ret_value
    End Function

    'return the full path & filename
    'WARNING: overwrites existing file of same name
    'display a message and return empty string on IO error
    Public Function write_csv(csv_name As String, fields As String, data As String) _
                              As String
        'full path + filename
        Dim file_name As String = dmt_working_path & csv_name

        'Write field headers & data to file
        Try
            Using sw As New StreamWriter(file_name)
                sw.WriteLine(fields)
                sw.Write(data)
            End Using
        Catch e As Exception
            MsgBox("The CSV file could not be writtern: " & e.Message)
            Return ""
        End Try

        'need to return the full path & filename to pass to DMT
        Return file_name
    End Function

    Public Sub dmt_log_event(prefix As String, msg As String)
        Dim file_path, file_name
        Dim log_date = DateTime.Now

        Dim log_msg As String
        Dim event_time = DateTime.Now
        log_msg = event_time.ToString("HH:mm:ss") & ": " & prefix & ": " & msg

        'create log directory - no filesystem changes will be made if it exists already
        file_path = dmt_working_path & "log\"
        Directory.CreateDirectory(file_path)

        file_name = file_path & log_date.ToString("yyyyMMdd") & "_dmtlog.txt"
        Try
            Using sw As StreamWriter = File.AppendText(file_name)
                sw.WriteLine(log_msg)
            End Using
        Catch e As Exception
            MsgBox("Couldn't write to DMT log file: " & e.Message)
        End Try
    End Sub

    'takes 1 argument: filename as fully-qualified filename of DMT-generated error file
    'returns: String containing user-actionable error conditions
    Public Function parse_dmt_error_log(ByVal filename As String) As String
        'regex groups to parse: `date` `related fields` `error message`
        '                group:  (1)          (2)             (3)
        Dim error_line_pattern As String = "^(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}) (.+?) (Table.*|Column.*)$"
        Dim error_regex As New Regex(error_line_pattern)

        Dim return_string As String = ""

        Try
            Using sr As New StreamReader(filename)
                Do while sr.Peek() >= 0
                    Dim log_line As String = sr.ReadLine()
                    Dim line_match As Match = error_regex.Match(log_line)

                    If Not line_match.Success Then
                        return_string = return_string & System.Environment.NewLine & _
                                        "Unrecognized error in log; please give this " & _
                                        "error message to your system administrator: " & log_line
                        Continue Do
                    End If

                    'parse the error
                    return_string = return_string & System.Environment.NewLine & _
                                    parse_error_line(log_line, line_match)
                Loop
            End Using
        Catch ex As FileNotFoundException
            'note: DMT removes existing error file of same name if rerun w/ no errors
            return_string = "DMT ran without errors"
        Catch ex As DirectoryNotFoundException
            return_string = "Log file directory not found: " & ex.Message
        Catch ex As IOException
            return_string = "DMT error file exists but cannot be read: " & ex.Message
        Catch ex As Exception
            'other possible exceptions for StreamReader:
            ' - `ArgumentException`/`ArgumentNullException` (empty/null filename, which shouldn't happen here)
            return_string = "Uncaught catastrophic failure. Contact your system " & _
                            "administrator with this error message: " & ex.Message
        End Try

        Return return_string
    End Function

    Private Function parse_error_line(ByVal log_line As String, ByRef line_match As Match) As String
        Dim return_string As String = ""

        Dim error_msg_pattern As String = "Table: (\w*) {0,1}Msg: (.+)$"
        Dim error_msg_regex As New Regex(error_msg_pattern)

        'the fields to notify the user about are conveniently space-delimited
        Dim error_fields As String()
        error_fields = line_match.Groups(2).Value.Split(New String() {" "}, StringSplitOptions.None)

        Dim msg_match As Match = error_msg_regex.Match(line_match.Groups(3).Value)
        If Not msg_match.Success AndAlso log_line.Contains("constrained to be unique") Then
            'if the above string is present in the DMT error message, then we're
            ' working with a BOM, and the row in the CSV corresponding to this
            ' error was rejected because the fields matching the compound primary
            ' key in Epicor match one already in the DB
            return_string = "Entry for part " & error_field(0) & ", revision " & _
                            error_field(1) & ", material " & error_field(2) & _
                            "already exists in the BOM in Epicor. It can " &_
                            "only be updated directly from Epicor ERP."
        ElseIf Not msg_match.Success
            return_string = "Unknown DMT error. Please show this error message to " & _
                            "your system administrator: " & log_line
        End If

        'if return string has already been populated above, return early
        ' (I just don't want the rest of this function in the Else block above)
        If Not String.IsNullOrEmpty(return_string) Then Return return_string

        Dim table As String = msg_match.Groups(1).Value
        Dim msg As String = msg_match.Groups(2).Value

        If String.IsNullOrEmpty(table) Then
            If String.Equals(msg, "Your software license does not allow this feature.") Then
                'pass: this error shouldn't appear anymore, but if it does it should be ignorable
                '      (eg, from an earlier version of the software)
            End If
        ElseIf String.Equals(table, "Part") Then
            If String.Equals(msg, "Part Number already exists.") Then
                return_string = "Part " & error_fields(0) & " has already been " & _
                                "exported into Epicor."
            Else
                return_string = "DMT has encountered an error exporting your part " & _
                                "that wasn't caught by validation: " & msg
            End If
        ElseIf String.Equals(table, "PartRev") Then
        ElseIf String.Equals(table, "ECOMtl") Then
        End If

        Return return_string
    End Function
End Class
