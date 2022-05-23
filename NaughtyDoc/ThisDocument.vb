Imports Microsoft.Win32

<ComClass()>
<Runtime.InteropServices.ComVisible(True)>
Public Class ThisDocument

    Public Sub ThisDocument_Startup() Handles Me.Startup
        Dim IsIndividuallyTrusted As Boolean
        IsIndividuallyTrusted = Security.IsDocumentTrusted(Me.Path & Me.Name)

        Dim IsTrustedByPath As Boolean
        IsTrustedByPath = Security.IsDocumentInTrustedLocation(Me.Path)
        If Not IsIndividuallyTrusted And Not IsTrustedByPath Then
            Dim msg As String
            msg = "This document is not trusted.  Do you want to trust it?"
            If MsgBox(msg, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                ' Offer the selection between folder or file trust
                If MsgBox("Do you want to trust this document as a folder [Yes], file [No], or cancel [Cancel]?", MsgBoxStyle.YesNoCancel) = MsgBoxResult.Yes Then
                    Security.TrustLocation(Me.Path, True)
                ElseIf MsgBoxResult.No Then
                    MsgBox("Functionality not implemented yet")
                ElseIf MsgBoxResult.Cancel Then
                    MsgBox("Okay.")
                End If
            End If
        Else
            MsgBox("This document is trusted.")
        End If
    End Sub

    Public Sub ThisDocument_Shutdown() Handles Me.Shutdown

    End Sub

    Public Sub CreateTable(ByVal firstColumnHeader As String,
                           ByVal secondColumnHeader As String)

        Me.Paragraphs(1).Range.InsertParagraphBefore()
        Dim table1 As Word.Table = Me.Tables.Add(Me.Paragraphs(1).Range, 2, 2)

        With table1
            .Style = "Table Professional"
            .Cell(1, 1).Range.Text = firstColumnHeader
            .Cell(1, 2).Range.Text = secondColumnHeader
        End With
    End Sub

End Class

Public NotInheritable Class Security
    'This class is used to validate security levels for the document in the trust center.
    ' This function checks to see if the individual document is trusted.
    Public Shared Function IsDocumentTrusted(documentAbsolutePath As String) As Boolean

        Dim localKey As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64)
        ' Search the registry for the TrustRecords values.
        Dim key As RegistryKey = localKey.OpenSubKey("Software\Microsoft\Office\16.0\Word\Security\Trusted Documents\TrustRecords")
        ' Get an array of the value names under the TrustRecords key.
        Dim valueNames As Object = key.GetValueNames()
        ' If valueNames is Nothing, exit the function and return false.
        If valueNames Is Nothing Then
            Return False
            ' Otherwise, proceed.
        Else
            ' Loop through the value names.
            For Each valueName As String In valueNames
                ' Since the path delimeter is a forward slash, replace it with a backslash.
                valueName = valueName.Replace("/", "\")
                ' If the value name matches the document path, return true.
                If valueName = documentAbsolutePath Then
                    Return True
                Else
                    ' If the value name does not match the document path, continue the loop.
                    Continue For
                End If
            Next
            ' If the loop completes, return false.
            Return False
        End If
        ' This should never be reached, throw an exception.
        Throw New Exception("The document path could not be found in the registry.")
    End Function

    Public Shared Function IsDocumentInTrustedLocation(documentPath As String) As Boolean
        Dim localKey As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64)
        Dim basekey As RegistryKey = localKey.OpenSubKey("Software\Microsoft\Office\16.0\Word\Security\Trusted Locations")
        Dim subkeyNames As Object = basekey.GetSubKeyNames()
        If subkeyNames Is Nothing Then
            Return False
        Else
            For Each subkeyName As String In subkeyNames
                Dim locationKey As RegistryKey = basekey.OpenSubKey(subkeyName)
                Dim locationPath As String = locationKey.GetValue("Path")
                Dim isRecursive As Boolean = locationKey.GetValue("AllowSubfolders")
                If isRecursive Then
                    If documentPath.StartsWith(locationPath) Then
                        Return True
                    End If
                Else
                    If documentPath = locationPath Then
                        Return True
                    End If
                End If
            Next
            Return False
        End If
    End Function

    Public Shared Function TrustLocation(path As String, Optional recursive As Boolean = False) As Boolean
        Dim localKey As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64)
        Dim basekey As RegistryKey = localKey.OpenSubKey("Software\Microsoft\Office\16.0\Word\Security\Trusted Locations", True)
        Dim subkeyNames As Object = basekey.GetSubKeyNames()
        ' Initialize the locationKeyStr variable to avoid undeclared variable error.
        Dim locationKeyStr As String

        If subkeyNames Is Nothing Then
            ' Assumes there is no current trusted locations.
            ' This means the key will start at Location0
            locationKeyStr = "Location0"
        Else
            ' Assumes there is at least one trusted location.
            ' This means the key will start one higher than the last trusted location.
            For Each subkeyName As String In subkeyNames
                Dim versionChar As Char
                Dim highestLocation As Integer = Convert.ToInt32(subkeyName.Substring(7))
                locationKeyStr = "Location" & (highestLocation + 1).ToString()
            Next
        End If

        ' Create the new trusted location key and set the values.
        Try
            Dim locationKey As RegistryKey = basekey.CreateSubKey(locationKeyStr)
            ' Convert values to appropriate RegistryValueKind types.
            ' The path is a REG_SZ, so use RegistryValueKind.String.
            locationKey.SetValue("Path", path, RegistryValueKind.String)
            ' The recursive value is a REG_DWORD, so use RegistryValueKind.DWord.
            ' Convert the boolean recursive variable to a compatible integer.
            Dim recursiveInt As Integer = If(recursive, 1, 0)
            locationKey.SetValue("AllowSubfolders", recursiveInt, RegistryValueKind.DWord)

            ' A Case to handle the response messages to the user.
            ' Responses include the path and whether or not the location is recursive.
            Select Case recursive
                Case True
                    MsgBox("Trusted location: " & path & " [ Recursive ]")
                Case False
                    MsgBox("Trusted location: " & path & " [ Non-Recursive ]")
            End Select
            Return True

        Catch
            ' A Case to handle the response messages to the user.
            ' Responses include the path and whether or not the location is recursive.
            Select Case recursive
                Case True
                    MsgBox("Failed to trust location: " & path & " [ Recursive ]")
                Case False
                    MsgBox("Failed to trust location: " & path & " [ Non-Recursive ]")
            End Select
            Return False
        End Try

    End Function
End Class