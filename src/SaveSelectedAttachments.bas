'====================================================================================
' Script Name : Save Selected Email Attachments
' Compatible  : Microsoft Outlook Desktop (All modern versions)
' Version     : 2.3 Performance Optimized
' Date        : 18-May-2026
'
' Features:
'   - Bulk-save attachments from selected Outlook emails
'   - Duplicate filename handling
'   - Illegal filename sanitization
'   - Inline/signature image skipping
'   - OLE object filtering
'   - MAX_PATH protection
'   - Performance optimized batch processing
'   - Outlook responsiveness protection
'   - Defensive error handling
'   - COM object cleanup
'
'====================================================================================

Option Explicit

Private Function SanitizeFileName(ByVal sName As String) As String
    Dim illegalChars As Variant
    Dim i As Integer
    
    illegalChars = Array("<", ">", ":", Chr(34), "/", "\", "|", "?", "*")

    For i = 0 To UBound(illegalChars)
        sName = Replace(sName, illegalChars(i), "_")
    Next i

    SanitizeFileName = sName
End Function

Public Sub SaveSelectedAttachments()

    On Error GoTo ErrorHandler

    Dim objSelection   As Outlook.Selection
    Dim objItem        As Object
    Dim objMail        As Outlook.MailItem
    Dim objAttachment  As Outlook.Attachment

    Dim ShellApp       As Object
    Dim FolderObj      As Object
    Dim SaveFolder     As String

    Dim FilePath       As String
    Dim SafeName       As String
    Dim BaseName       As String
    Dim Extension      As String
    Dim Counter        As Long

    Dim SavedCount     As Long
    Dim MailCount      As Long
    Dim SkippedCount   As Long

    Const PR_ATTACH_FLAGS As String = "http://schemas.microsoft.com/mapi/proptag/0x37140003"
    Dim lngFlags As Long

    Set objSelection = Application.ActiveExplorer.Selection

    If objSelection.Count = 0 Then
        MsgBox "Please select emails first.", vbExclamation
        Exit Sub
    End If

    Set ShellApp = CreateObject("Shell.Application")
    Set FolderObj = ShellApp.BrowseForFolder(0, "Select Folder To Save Attachments", 0)

    If FolderObj Is Nothing Then
        MsgBox "Operation cancelled.", vbInformation
        Set ShellApp = Nothing
        Exit Sub
    End If

    SaveFolder = FolderObj.self.Path

    If Right(SaveFolder, 1) <> "\" Then
        SaveFolder = SaveFolder & "\"
    End If

    For Each objItem In objSelection

        If TypeName(objItem) = "MailItem" Then

            Set objMail = objItem
            MailCount = MailCount + 1

            If MailCount Mod 25 = 0 Then DoEvents

            If objMail.Attachments.Count > 0 Then

                For Each objAttachment In objMail.Attachments

                    lngFlags = 0

                    On Error Resume Next
                    lngFlags = objAttachment.PropertyAccessor.GetProperty(PR_ATTACH_FLAGS)

                    If Err.Number <> 0 Then Err.Clear

                    On Error GoTo ErrorHandler

                    If lngFlags = 4 Then
                        SkippedCount = SkippedCount + 1
                        GoTo NextAttachment
                    End If

                    If objAttachment.Type <> olByValue Then
                        SkippedCount = SkippedCount + 1
                        GoTo NextAttachment
                    End If

                    SafeName = SanitizeFileName(objAttachment.FileName)
                    FilePath = SaveFolder & SafeName

                    If Dir(FilePath) <> "" Then

                        If InStrRev(SafeName, ".") > 0 Then
                            BaseName = Left(SafeName, InStrRev(SafeName, ".") - 1)
                            Extension = Mid(SafeName, InStrRev(SafeName, "."))
                        Else
                            BaseName = SafeName
                            Extension = ""
                        End If

                        Counter = 1

                        Do While Dir(SaveFolder & BaseName & "_" & Counter & Extension) <> ""
                            Counter = Counter + 1
                        Loop

                        SafeName = BaseName & "_" & Counter & Extension
                        FilePath = SaveFolder & SafeName

                    End If

                    If Len(FilePath) > 259 Then
                        SkippedCount = SkippedCount + 1
                        GoTo NextAttachment
                    End If

                    objAttachment.SaveAsFile FilePath

                    SavedCount = SavedCount + 1

NextAttachment:

                Next objAttachment

            End If

        End If

    Next objItem

    Set objSelection = Nothing
    Set objItem = Nothing
    Set objMail = Nothing
    Set objAttachment = Nothing
    Set FolderObj = Nothing
    Set ShellApp = Nothing

    MsgBox _
        "Completed Successfully!" & vbCrLf & vbCrLf & _
        "Emails Processed     : " & MailCount & vbCrLf & _
        "Attachments Saved    : " & SavedCount & vbCrLf & _
        "Skipped (inline/OLE) : " & SkippedCount & vbCrLf & _
        "Saved To             : " & SaveFolder, _
        vbInformation

    Exit Sub

ErrorHandler:

    MsgBox _
        "Error Occurred:" & vbCrLf & _
        Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
        "Progress before error:" & vbCrLf & _
        "  Emails Processed  : " & MailCount & vbCrLf & _
        "  Attachments Saved : " & SavedCount, _
        vbCritical

    Set ShellApp = Nothing
    Set FolderObj = Nothing

End Sub
