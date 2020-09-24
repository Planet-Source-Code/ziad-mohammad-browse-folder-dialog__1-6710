<div align="center">

## Browse Folder Dialog


</div>

### Description

Have ever wondered if there is an ActiveX object that make you browse for a folder. This API functions calls make the browse dialog
 
### More Info
 
Start a new Project and Add a command button on the form named command1

This API function calls display the structure of your computer and allow the use to select a folder


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ziad Mohammad](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ziad-mohammad.md)
**Level**          |Advanced
**User Rating**    |5.0 (60 globes from 12 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ziad-mohammad-browse-folder-dialog__1-6710/archive/master.zip)

### API Declarations

```
Private Const BIF_RETURNONLYFSDIRS = 1
   Private Const BIF_DONTGOBELOWDOMAIN = 2
   Private Const BIF_BROWSEFORCOMPUTER = &H1000
   Private Const MAX_PATH = 260
   Private Declare Function SHBrowseForFolder Lib "shell32" _
                    (lpbi As BrowseInfo) As Long
   Private Declare Function SHGetPathFromIDList Lib "shell32" _
                    (ByVal pidList As Long, _
                    ByVal lpBuffer As String) As Long
   Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                    (ByVal lpString1 As String, ByVal _
                    lpString2 As String) As Long
   Private Type BrowseInfo
     hwndOwner   As Long
     pIDLRoot    As Long
     pszDisplayName As Long
     lpszTitle   As Long
     ulFlags    As Long
     lpfnCallback  As Long
     lParam     As Long
     iImage     As Long
    End Type
```


### Source Code

```
Private Sub Command1_Click()
   'Opens a Treeview control that displays the directories in a computer
  Dim lpIDList As Long
  Dim sBuffer As String
  Dim szTitle As String
  Dim tBrowseInfo As BrowseInfo
 szTitle = "This is the title"
 With tBrowseInfo
  .hWndOwner = Me.hWnd
  .lpszTitle = lstrcat(szTitle, "")
  .ulFlags = BIF_RETURNONLYFSDIRS_
  +BIF_DONTGOBELOWDOMAIN
 End With
 lpIDList = SHBrowseForFolder(tBrowseInfo)
 If (lpIDList) Then
      sBuffer = Space(MAX_PATH)
      SHGetPathFromIDList lpIDList, sBuffer
      sBuffer = Left(sBuffer, InStr
      (sBuffer, vbNullChar) - 1)
      MsgBox sBuffer
 End If
End Sub
```

