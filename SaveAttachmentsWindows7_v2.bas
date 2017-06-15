Attribute VB_Name = "Module2"
Dim DefFolder As String

' SaveAttachments()
'   Subroutine saves all attachments in the current mail message
'   to a folder specified by the user, deletes the attachments
'   from the message, and inserts a comment into the message
'   listing the files that have been saved.
'
' Notes:
'   Running this macro will change message format to Rich Text
'   and will remove all formatting.  There may be a way around
'   this for HTML messages using the HTMLBody property.
'
' 1/3/03 P. Tatian
' 1/9/03 Updated to add file date/time to message asking about
'        replacing existing files
'
' 6/18/12 Changed CreateObject("Outlook.Application") to CreateObject("Outlook.Application.14")
'         to solve problem with "Run-time error '429': ActiveX component can't create object"
'         message when running under Windows 7. Ref: http://support.microsoft.com/kb/828550
'
' 5/6/14  Version 2. Insert hyperlinks into mail messages. Converts messages to HTML.
'
' 6/15/17 Updated for Outlook 2013 or later
Sub SaveAttachments()

  Dim myInspector As Outlook.Inspector
  Dim myItem As Object
  Dim AttCount As Integer, I As Integer, Count As Integer
  Dim SaveFolder As String, InsertMsg As String, PathName As String
  Dim MsgStr As String
  Dim AttSaved As Boolean
  
  Set myInspector = Application.ActiveInspector

  On Error Resume Next
  Set myItem = myInspector.CurrentItem
  
  ' Check if current item is a mail message
  
  If myItem.Class <> olMail Then
    MsgBox "You must open a mail message before running this macro.", _
           vbExclamation, _
           "Not a mail message!"
    Exit Sub
  End If

  ' Check if mail message has been saved

  If Not myItem.Saved Then
    MsgBox "Mail message must be saved before running this macro.", _
           vbExclamation, _
           "Message not saved!"
    Exit Sub
  End If

  ' Determine number of attachments.  If none, cancel macro.

  AttCount = myItem.Attachments.Count

  If AttCount = 0 Then
    MsgBox "The current mail message has no attachments.", _
           vbExclamation, _
           "No attachments!"
    Exit Sub
  End If

  ' Ask user for folder where attachments should be saved.
  ' SaveFolder is empty if user presses Cancel button on dialog.

  SaveFolder = Trim(GetFolderName("Enter folder where attachments should be saved."))
  
  If SaveFolder = "" Then
    Exit Sub
  End If
  
  ' Check for trailing "\" at end of folder name
  
  If Mid(SaveFolder, Len(SaveFolder), 1) <> "\" Then
    SaveFolder = SaveFolder & "\"
  End If
  
  ' Initialize insertion list message
    
  InsertMsg = "<br />" & "<br />" & "[Attachments saved: "
  
  ' Initialize attachment pointer (I) and attachment saved flag
  I = 1
  AttSaved = False
  
  ' Process each attachment
  
  For Count = 1 To AttCount
  
    ' Get complete pathname for file to be saved
  
    PathName = SaveFolder & myItem.Attachments(I).FileName
  
    ' Check whether file already exists
    
    If FileExists(PathName) Then
      
      ' File exists, ask user if it should be replaced
      
      MsgStr = "File " & PathName & " (" & FileDate(PathName) & _
               ")" & vbCrLf & "already exists." & vbCrLf & vbCrLf _
               & "Do you want to replace it?"
      
      Ans = MsgBox(MsgStr, _
                   vbYesNoCancel + vbQuestion + vbDefaultButton2, _
                   "File exists!")
                   
      ' Process user answer
                   
      Select Case Ans
      
        Case vbYes
          ' Yes: Just continue normally
                    
        Case vbNo
          ' No:  Skip attachment
          I = I + 1
          GoTo NextCount
        
        Case vbCancel
          ' Cancel:  Exit procedure completely
          Exit Sub
        
      End Select
        
    End If
  
    ' Save attachment and set flag
    myItem.Attachments(I).SaveAsFile (PathName)
    AttSaved = True
    
    ' Delete attachment
    myItem.Attachments(I).Delete
    
    ' Add file name to insertion message
    InsertMsg = InsertMsg & "<br />" & "&nbsp;&nbsp;" & "<a href='" & PathName & "'>" & PathName & "</a>"
  
NextCount:
  
  Next Count
  
  ' Complete insertion message
  
  InsertMsg = InsertMsg & " ]" & "<br />"
  
  ' If an attachment was saved, insert list of saved files in message
  
  If AttSaved Then
    myItem.HTMLBody = myItem.HTMLBody & InsertMsg
  End If

End Sub

' DeleteAttachments()
'   Subroutine deletes all attachments in the current mail message
'   and inserts a comment into the message
'   listing the files that have been deleted.
'
' Notes:
'   Running this macro will change message format to Rich Text
'   and will remove all formatting.  There may be a way around
'   this for HTML messages using the HTMLBody property.
'
' 2/11/03 P. Tatian
' 9/1/04  Added confirmation before deleting attachments;
'         removed final message
'
' 6/18/12 Changed CreateObject("Outlook.Application") to CreateObject("Outlook.Application.14")
'         to solve problem with "Run-time error '429': ActiveX component can't create object"
'         message when running under Windows 7. Ref: http://support.microsoft.com/kb/828550
'
' 5/6/14  Version 2. Retain formatting for HTML messages. Other message types are converted to rich text.
Sub DeleteAttachments()

  Dim myInspector As Outlook.Inspector
  Dim myItem As Object
  Dim AttCount As Integer, I As Integer, Count As Integer
  Dim SaveFolder As String, InsertMsg As String, PathName As String
  Dim MsgStr As String
  Dim AttSaved As Boolean
  
  Set myInspector = Application.ActiveInspector

  On Error Resume Next
  Set myItem = myInspector.CurrentItem
  
  ' Check if current item is a mail message
  
  If myItem.Class <> olMail Then
    MsgBox "You must open a mail message before running this macro.", _
           vbExclamation, _
           "Not a mail message!"
    Exit Sub
  End If

  ' Check if mail message has been saved

  If Not myItem.Saved Then
    MsgBox "Mail message must be saved before running this macro.", _
           vbExclamation, _
           "Message not saved!"
    Exit Sub
  End If

  ' Determine number of attachments.  If none, cancel macro.

  AttCount = myItem.Attachments.Count

  If AttCount = 0 Then
    MsgBox "The current mail message has no attachments.", _
           vbExclamation, _
           "No attachments!"
    Exit Sub
  End If
  
  ' Check to make sure user really wants to delete attachments
  
  Ans = MsgBox("Do you really want to delete all attachments?", _
                   vbYesNo + vbQuestion + vbDefaultButton2, _
                   "Confirm delete")
                   
  If Ans = vbNo Then
    Exit Sub
  End If
  
  ' Determine message format and set line break and indent character strings
  If myItem.BodyFormat = olFormatHTML Then
    ' HTML message
    LineBreak = "<br />"
    Indent = "&nbsp;&nbsp;"
  Else
    ' Other message formats
    LineBreak = vbCrLf
    Indent = "    "
  End If

  ' Initialize insertion list message
    
  InsertMsg = LineBreak & LineBreak & "[Attachments deleted: "
  
  ' Delete each attachment
  
  For Count = 1 To AttCount
  
    ' Add file name to insertion message
    InsertMsg = InsertMsg & LineBreak & Indent & myItem.Attachments(1).FileName
  
    ' Delete attachment
    myItem.Attachments(1).Delete
    
  Next Count
  
  ' Complete insertion message
  
  InsertMsg = InsertMsg & " ]" & LineBreak
  
  ' Insert list of deleted files in message
  
  If myItem.BodyFormat = olFormatHTML Then
    ' HTML message
    myItem.HTMLBody = myItem.HTMLBody & InsertMsg
  Else
    ' Other message formats
    myItem.Body = myItem.Body & InsertMsg
  End If

  'MsgBox "Attachments deleted.", _
           vbExclamation, _
           "DeleteAttachments macro"

End Sub

' GetFolderName()
'   Prompts user for a name of a valid folder.  If the folder provided
'   does not exist, user is prompted to reenter it.
'
Function GetFolderName(Message As String)

  Dim Title, Default, MyValue, fs
  
  Title = "Enter folder"    ' Set title.
  
  If DefFolder = "" Then
    Default = "D:\"    ' Set default.
  Else
    Default = DefFolder
  End If

  Do

  ' Display message, title, and default value.
  MyValue = InputBox(Message, Title, Default)
  
  If MyValue = "" Then
    GetFolderName = ""
    Exit Do
  End If
  
  Set fs = CreateObject("Scripting.FileSystemObject")
  
  If fs.FolderExists(MyValue) Then
    DefFolder = MyValue
    GetFolderName = MyValue
    Exit Do
  Else
    MsgBox "The folder " & MyValue & " does not exist.", _
           vbExclamation + vbOKOnly, _
           "Folder does not exist!"
    Default = MyValue
  End If
  
  Loop
  
End Function

' FileExists()
'    Returns True if file specified in PathName exists,
'    False if it does not.
'
Function FileExists(PathName As String)

  Dim fs

  Set fs = CreateObject("Scripting.FileSystemObject")
  
  FileExists = fs.FileExists(PathName)

End Function

' FileDate()
'    Returns the date that file in PathName was last modified,
'    or an empty string if the file does not exist.
'
Function FileDate(PathName As String)

  Dim fs, f

  Set fs = CreateObject("Scripting.FileSystemObject")
  Set f = fs.GetFile(PathName)

  If fs.FileExists(PathName) Then
    FileDate = f.DateLastModified
  Else
    FileDate = ""
  End If

End Function

