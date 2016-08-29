Imports System.Collections.Generic

Public Module EpicorOps
    'create an ArrayList of appropriate options from CSV `f`
    'the options selected from the CSV match the part type specified
    '`f` is just the CSV filename, the path is retrieved from the DMT class
    'the CSV should in turn be populated by DMT's dmt_export() method
    Public Function fetch_list_values(ByVal f As String, _
                                      ByVal dmt_working_path As String, _
                                      ByVal part_type As String) As ArrayList
        Dim file_name As String = dmt_working_path & "ref\" & f
        Dim option_list As New ArrayList()

        Using csv_reader As New FileIO.TextFieldParser(file_name)
            csv_reader.TextFieldType = FileIO.FieldType.Delimited
            csv_reader.SetDelimiters(",")

            Dim current_row As String()
            Dim first_line As Boolean = True
            While Not csv_reader.EndOfData
                Try
                    current_row = csv_reader.ReadFields()
                Catch ex As FileIO.MalformedLineException
                    Debug.Write("CSV contained invalid line:" & ex.Message)
                End Try

                If first_line Then
                    'skip the header row
                    first_line = False
                ElseIf current_row(2).Equals(part_type) Then
                    'the first field contains the human-readable description
                    ' that will populate the options list, and the third contains
                    ' the part type to determine whether it should be added
                    option_list.Add(current_row(0))
                End If
            End While
        End Using

        Return option_list
    End Function

    'Return a dictionary mapping the description in the parameter to the DB
    ' friendly ID expected by Epicor
    Function fetch_list_mappings(ByVal f As String, _
                                 ByVal dmt_working_path As String) _
                                 As Dictionary(Of String, String)
        Dim file_name As String = dmt_working_path & "ref\" & f
        Dim mapping As New Dictionary(Of String, String)

        Using csv_reader As New FileIO.TextFieldParser(file_name)
            csv_reader.TextFieldType = FileIO.FieldType.Delimited
            csv_reader.SetDelimiters(",")

            Dim current_row As String()
            Dim first_line As Boolean = True
            While Not csv_reader.EndOfData
                Try
                    current_row = csv_reader.ReadFields()
                Catch ex As FileIO.MalformedLineException
                    Debug.Write("CSV contained invalid line:" & ex.Message)
                End Try

                If first_line Then
                    'skip headers
                    first_line = False
                Else
                    mapping.Add(current_row(0), current_row(1))
                End If
            End While
        End Using

        Return mapping
    End Function
End Module
