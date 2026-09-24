'====================================================================================
' Script Name : Save Actual Files Only (Production Grade V3.1)
' Compatible  : Microsoft Outlook Desktop (All modern versions)
' Description : Extracts likely user-visible file attachments while filtering 
'               common inline/signature images.
'====================================================================================

Option Explicit

Public Sub SaveActualFilesOnly()
    '------------------------------------------------------------
    ' 1. Variable Declarations (All moved to the top)
    '------------------------------------------------------------
    Dim objSelection As Outlook.Selection
    Dim objItem As Object, objMail As Outlook.MailItem
    Dim objAttachment As Outlook.Attachment
    
    Dim ShellApp As Object, FolderObj As Object
    Dim objFSO As Object
    
    Dim SaveFolder As String, FilePath As String
    Dim BaseName As String, Ext As String, CleanName As String
    
    Dim MailCount As Long, SavedCount As Long
    Dim SkippedCount As Long, FailedCount As Long, Counter As Long
    
    Dim vCID As Variant, vHidden As Variant
    Dim isLikelyInline As Boolean
    Dim strExtCheck As String
    Dim ReportMsg As String

    '------------------------------------------------------------
    ' 2. Check Explorer State & Initialize Picker
    '------------------------------------------------------------
    If Application.ActiveExplorer Is Nothing Then
        MsgBox "No active Outlook window found. Please select an email.", vbExclamation
        Exit Sub
    End If

    Set objSelection = Application.ActiveExplorer.Selection
    If objSelection.Count = 0 Then
        MsgBox "Please select at least one email.", vbExclamation
        Exit Sub
    End If

    Set ShellApp = CreateObject("Shell.Application")
    Set FolderObj = ShellApp.BrowseForFolder(0, "Select Folder To Save Attachments", 0)

    If FolderObj Is Nothing Then Exit Sub
    
    ' Cleaner folder path construction (No IIf used)
    SaveFolder = FolderObj.Self.Path
    If Right$(SaveFolder, 1) <> "\" Then
        SaveFolder = SaveFolder & "\"
    End If
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    '------------------------------------------------------------
    ' 3. Process Emails
    '------------------------------------------------------------
    For Each objItem In objSelection
        If TypeName(objItem) = "MailItem" Then
            Set objMail = objItem
            MailCount = MailCount + 1
            
            For Each objAttachment In objMail.Attachments
                
                ' =======================================================
                ' TOTAL ISOLATION: Shield the entire attachment iteration
                ' =======================================================
                On Error Resume Next
                Err.Clear
                
                ' olByValue (1) targets actual files/documents only, ignoring attached .msg files
                If objAttachment.Type = olByValue Then 
                    
                    isLikelyInline = False
                    vCID = Empty
                    vHidden = Empty
                    
                    ' Safely query MAPI properties
                    vCID = objAttachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E")
                    Err.Clear
                    
                    vHidden = objAttachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x7FFE000B")
                    Err.Clear
                    
                    ' Evaluate Heuristics
                    If VarType(vHidden) = vbBoolean Then
                        If vHidden = True Then isLikelyInline = True
                    End If
                    
                    If Not isLikelyInline And VarType(vCID) = vbString Then
                        If vCID <> "" Then
                            strExtCheck = LCase(objFSO.GetExtensionName(objAttachment.FileName))
                            If strExtCheck = "png" Or strExtCheck = "jpg" Or strExtCheck = "jpeg" Or strExtCheck = "gif" Then
                                isLikelyInline = True
                            End If
                        End If
                    End If
                    
                    ' Save the file if it passes the heuristic filter
                    If Not isLikelyInline Then
                        
                        CleanName = CleanFilename(objAttachment.FileName)
                        BaseName = objFSO.GetBaseName(CleanName)
                        Ext = objFSO.GetExtensionName(CleanName)
                        If Ext <> "" Then Ext = "." & Ext
                        
                        FilePath = SaveFolder & CleanName
                        Counter = 1
                        
                        ' FSO File Exists Check
                        Do While objFSO.FileExists(FilePath)
                            FilePath = SaveFolder & BaseName & "_" & Counter & Ext
                            Counter = Counter + 1
                        Loop
                        
                        objAttachment.SaveAsFile FilePath
                        
                        ' Tally successes or failures for this specific attachment
                        If Err.Number = 0 Then
                            SavedCount = SavedCount + 1
                        Else
                            FailedCount = FailedCount + 1
                        End If
                        
                    Else
                        SkippedCount = SkippedCount + 1
                    End If
                    
                End If
                
                ' Reset error handling before moving to the next attachment
                Err.Clear
                On Error GoTo 0
                ' =======================================================
                
            Next objAttachment
        End If
    Next objItem

    ' Clean up objects
    Set objFSO = Nothing
    Set ShellApp = Nothing
    Set FolderObj = Nothing

    '------------------------------------------------------------
    ' 4. Granular Reporting
    '------------------------------------------------------------
    ReportMsg = "Batch Processing Complete!" & vbCrLf & vbCrLf & _
                "Emails Processed: " & MailCount & vbCrLf & _
                "Files Saved: " & SavedCount & vbCrLf & _
                "Inline/Logos Skipped: " & SkippedCount
                
    If FailedCount > 0 Then
        ReportMsg = ReportMsg & vbCrLf & "Failed to Save: " & FailedCount & " (Check permissions or corrupted files)"
    End If

    MsgBox ReportMsg, IIf(FailedCount > 0, vbExclamation, vbInformation), "Status Report"

End Sub

'------------------------------------------------------------
' HELPER FUNCTION: Comprehensive Windows Filename Sanitizer
'------------------------------------------------------------
Private Function CleanFilename(ByVal strName As String) As String
    Dim invalidChars As Variant
    Dim reservedNames As Variant
    Dim i As Integer
    Dim fso As Object
    Dim baseN As String
    
    ' 1. Strip illegal characters
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(invalidChars) To UBound(invalidChars)
        strName = Replace(strName, invalidChars(i), "_")
    Next i
    
    ' 2. Strip trailing spaces and periods
    strName = Trim(strName)
    Do While Right$(strName, 1) = "." Or Right$(strName, 1) = " "
        strName = Left$(strName, Len(strName) - 1)
    Loop
    
    ' 3. Handle DOS Reserved Names
    Set fso = CreateObject("Scripting.FileSystemObject")
    baseN = UCase(fso.GetBaseName(strName))
    
    reservedNames = Array("CON", "PRN", "AUX", "NUL", _
                          "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", _
                          "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9")
                          
    For i = LBound(reservedNames) To UBound(reservedNames)
        If baseN = reservedNames(i) Then
            strName = "_" & strName ' Prepend underscore to bypass OS restriction
            Exit For
        End If
    Next i
    
    ' 4. Fallback for completely empty names
    If Trim(strName) = "" Then strName = "Unnamed_Attachment"
    
    Set fso = Nothing
    CleanFilename = strName
End Function
