' <IsStraightVb>True</IsStraightVb>
Imports System.Text.RegularExpressions
Imports System.Diagnostics
Imports System.IO

'class representing the DMT - stores config info, working paths, etc.
'responsible for setting up and calling the process
Public Class DMT
    Public Shared dmt_loc As String = "C:\Epicor\ERP10.1Client\Client\DMT.exe"
    Public Shared dmt_working_path As String = "I:\Cadd\_Epicor\"

    Private username As String
    Private password As String
    Private configfile As String
    Private connection As String
    Private dmt_base_args As String

    Public dmt_parsed_log As String

    'constructor
    Public Sub New()
        'TODO: change in production to DMT user/password/environment
        username = "DMT_USERNAME"
        password = "DMT_PASSWORD"
        configfile = "EpicorPilot10"
        connection = "net.tcp://CHERRY/EpicorPilot10"
        dmt_base_args = "-NoUI -User=" & username & " -Pass=" & password & " -ConnectionURL=""" & _
                        connection & """ -ConfigValue=""" & configfile & """"

        dmt_parsed_log = ""
    End Sub

    'Run the DMT to import the specified CSV into Epicor
    'Return values > 0 are passed along from log parser; < 0 signifies DMT exec timeout
    '`update_on` controls whether DMT only attempts to add the record as a new entry,
    ' or update any existing ones
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

        Dim return_code As Integer = exec_dmt(psi, csv, msg_succ)
        If return_code > 0 Then
            return_code = parse_dmt_error_log(filename)
        End If

        Return return_code
    End Function

    'use the DMT to export data from Epicor based on existing BAQs
    'the results of the queries is stored in the paired CSV files for later reading
    'TODO: pass along the return code from the DMT (-1 if it timed out)
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

        For Each kvp As KeyValuePair(Of String, String) In query_map
            psi.Arguments = dmt_base_args & " -Export -BAQ=""" & kvp.Key
            psi.Arguments = psi.Arguments & """ -Target=""" & export_path & kvp.Value & """"

            msg_succ = "Successfully exported CSV from Epicor"
            exec_dmt(psi, kvp.Key, msg_succ)
        Next
    End Sub

    'return -1 if DMT times out, otherwise pass on DMT's return value
    '0 = success
    '1 = error
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

    'write a csv with filename `csv_name` in the DMT's working directory, with
    ' `fields` as the first row and `data` as the, uh, data
    'display a message and return empty string on IO error
    'WARNING: overwrites existing file of same name
    'return the full path & filename
    Public Function write_csv(csv_name As String, fields As String, _
                              data As String) As String
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

    'write the specified `msg` to the module's logfile,
    ' timestamped and marked with the caller's `prefix`
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

    'takes 1 argument: fully-qualified filename of CSV passed to DMT
    'parses the DMT error log corresponding to the specified file and stores more
    ' readable versions of the errors in the DMT object's `dmt_parsed_log` member variable
    'returns: Integer corresponding to DMT return state
    '         - 0 = no errors
    '         - 1 = at least one error parsed in log file
    '         - 2 = I/O error
    '         - 3 = other unhandled error
    Private Function parse_dmt_error_log(ByVal filename As String) As Integer
        'DMT error file lines have a consistent formula
        'regex groups to parse: `date` `related fields` `error message`
        '                group:  (1)          (2)             (3)
        Dim error_line_pattern As String = "^(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})\s*(.+?)\s*(Table.*|Column.*)$"
        Dim error_regex As New Regex(error_line_pattern)

        Dim return_code As Integer = 0
        Dim return_string As String = ""

        Dim error_log_suffix As String = ".Errors.txt"

        Try
            Using sr As New StreamReader(filename & error_log_suffix)
                Do While sr.Peek() >= 0
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

                If Not String.IsNullOrEmpty(return_string) Then return_code = 1
            End Using
        Catch ex As FileNotFoundException
            'note: DMT removes existing error file of same name if rerun w/ no errors
            return_code = 0
            return_string = "DMT ran without errors"
        Catch ex As DirectoryNotFoundException
            return_code = 2
            return_string = "Log file directory not found: " & ex.Message
        Catch ex As IOException
            return_code = 2
            return_string = return_string & "DMT error file exists but cannot be read: " & ex.Message
        Catch ex As Exception
            'other possible exceptions for StreamReader:
            ' - `ArgumentException`/`ArgumentNullException` (empty/null filename, which shouldn't happen here)
            return_code = 3
            return_string = "Uncaught catastrophic failure. Contact your system " & _
                            "administrator with this error message: " & ex.Message
        End Try

        'the compiled message is stored in a member variable of the DMT object
        dmt_parsed_log = return_string
        Return return_code
    End Function

    'helper function to parse individual lines of a DMT error log
    'returns string containing more user-actionable error message
    Private Function parse_error_line(ByVal log_line As String, ByRef line_match As Match) As String
        Dim return_string As String = ""

        'the most common pattern that I've seen
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
            return_string = "Entry for part " & error_fields(0) & ", revision " & _
                            error_fields(1) & ", material " & error_fields(2) & _
                            " already exists in the BOM in Epicor. It can " & _
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

        Dim catchall_error As String = "DMT has encountered an unexpected error. " & _
                "Please forward this message to your system administrator: "
        return_string = catchall_error & msg

        'handle error messages originating from different tables in Epicor
        Select Case table
            Case ""
                If String.Equals(msg, "Your software license does not allow this feature.") Then
                    'pass: this error shouldn't appear anymore, but if it does it should be ignorable
                    '      (eg, from an earlier version of the software)
                End If
            Case "Part"
                If String.Equals(msg, "Part Number already exists.") Then
                    return_string = "Part " & error_fields(0) & " has already been " & _
                                    "exported into Epicor."
                End If
            Case "PartRev"
                If String.Equals(msg, "Record not available.") Then
                    If String.IsNullOrEmpty(error_fields(1)) Then
                        return_string = "You need to specify a revision number to export " & _
                                        "this Bill of Materials."
                    Else
                        return_string = "The part or revision for this Bill of Materials " & _
                                        "hasn't been entered into Epicor inventory yet."
                    End If
                End If
            Case "ECOMtl"
                If msg.Contains("Invalid Component Part Number") Then
                    return_string = "Material " & error_fields(2) & " referenced in this " & _
                                    "Bill of Materials hasn't been entered into Epicor " & _
                                    "inventory yet."
                End If
        End Select

        Return return_string
    End Function

    'pop up a message box to provide the user feedback based on
    ' the error code and parsed message
    Public Sub check_errors(ByVal ret_value As Integer, ByVal export_type As String)
        If ret_value = -1 Then
            MsgBox("Error: DMT timed out when processing the " & export_type & ". Aborting...")
        Else
            MsgBox("Aborting due to the following errors DMT experienced while processing the " & _
                export_type & ": "& System.Environment.NewLine & dmt_parsed_log)
        End If
    End Sub
End Class
