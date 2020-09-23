<div align="center">

## Kill Folder function with persistence

<img src="PIC2009107733503497.jpg">
</div>

### Description

This is a Kill Folder function with persistence ... It will remove all sub-folders and files and then optionally delete the specified folder ... I found when removing all the files in the temp folder that some locked files would fail and cause it to not continue with the rest of the files ... This function will continue to remove all unlocked files, even after finding locked files. However, if locked files are found, the parent folder will also not get removed.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rde](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rde.md)
**Level**          |Beginner
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rde-kill-folder-function-with-persistence__1-72521/archive/master.zip)





### Source Code


<tt>
<p nowrap><b>
<nobr>
<font color="#000099">
Option Explicit <br />
&#160; <br />
Private Declare Function</font> <font color="#660000">GetAttributes</font> <font color="#000099">Lib</font> "kernel32" <font color="#000099">Alias</font> "GetFileAttributesA" <font color="#000099">(ByVal</font> <font color="#660000">lpSpec</font> <font color="#000099">As String) As Long <br />
Private Declare Function</font> <font color="#660000">SetAttributes</font> <font color="#000099">Lib</font> "kernel32" <font color="#000099">Alias</font> "SetFileAttributesA" <font color="#000099">(ByVal</font> <font color="#660000">lpSpec</font> <font color="#000099">As String, ByVal</font> <font color="#660000">dwAttributes</font> <font color="#000099">As Long) As Long <br />
&#160; <br />
Private Const</font> <font color="#660000">DIR_SEP</font> <font color="#000099">As String</font> = "\" <br />
<font color="#000099">Private Const</font> <font color="#660000">INVALID_FILE_ATTRIBUTES</font> = (-1) <br />
&#160; <br />
<font color="#000099">
&#160;'-----------------------------------------------------</font> <br />
&#160; <br />
<font color="#006600">
&#160;' This is a Kill Folder function with persistence. <br />
&#160; <br />
&#160;' It will remove all sub-folders and files and then <br />
&#160;' optionally delete the specified folder.<br />
&#160; <br />
&#160;' I found when removing all the files in the temp folder <br />
&#160;' that some locked files would fail and cause it to not<br />
&#160;' continue with the rest of the files. <br />
&#160; <br />
&#160;' This function will continue to remove all unlocked files, <br />
&#160;' even after finding locked files. However, if locked files <br />
&#160;' are found, the parent folder will also not get removed.</font> <br />
&#160; <br />
<font color="#000099">
&#160;'-----------------------------------------------------<br />
</font>
&#160; <br />
<font color="#000099">Public Function</font> <font color="#660000">AddBackslash(sPath</font> <font color="#000099">As String) As String <br />
 &#160; If</font> <font color="#660000">Right$(sPath, 1&) = DIR_SEP</font> <font color="#000099">Then</font> <br />
 &#160; &#160; &#160;<font color="#660000">AddBackslash = sPath</font> <br />
 &#160; <font color="#000099">Else</font> <br />
 &#160; &#160; &#160;<font color="#660000">AddBackslash = sPath & DIR_SEP</font> <br />
 &#160; <font color="#000099">End If <br />
End Function <br />
&#160; <br />
&#160;'-----------------------------------------------------<br />
&#160; <br />
Public Function</font> <font color="#660000">FolderExists(sPath</font> <font color="#000099">As String) As Boolean <br />
 &#160; Dim</font> <font color="#660000">Attribs</font> <font color="#000099">As Long</font> <br />
 &#160; <font color="#660000">Attribs = GetAttributes(sPath)</font> <br />
 &#160; <font color="#000099">If Not</font> <font color="#660000">(Attribs = INVALID_FILE_ATTRIBUTES)</font> <font color="#000099">Then</font> <br />
 &#160; &#160; &#160;<font color="#660000">FolderExists = ((Attribs</font> <font color="#000099">And</font> <font color="#660000">vbDirectory) = vbDirectory)</font> <br />
 &#160; <font color="#000099">End If <br />
End Function <br />
&#160; <br />
&#160;'-----------------------------------------------------<br />
&#160; <br />
Public Function</font> <font color="#660000">KillFolder(sSpec</font> <font color="#000099">As String, Optional ByVal</font> <font color="#660000">bJustEmptyDontRemove</font> <font color="#000099">As Boolean) As Boolean <br />
 &#160; Dim</font> <font color="#660000">sRoot</font> <font color="#000099">As String,</font> <font color="#660000">sDir</font> <font color="#000099">As String,</font> <font color="#660000">sFile</font> <font color="#000099">As String <br />
 &#160; Dim</font> <font color="#660000">iCnt</font> <font color="#000099">As Long,</font> <font color="#660000">iIdx</font> <font color="#000099">As Long <br />
&#160; <br />
 &#160; If Not</font> <font color="#660000">FolderExists(sSpec)</font> <font color="#000099">Then Exit Function</font> <br />
&#160; <br />
 &#160; <font color="#006600">' Add trailing backslash if missing</font> <br />
 &#160; <font color="#660000">sRoot = AddBackslash(sSpec) <br />
 &#160; iCnt</font> = 2& <font color="#006600">'.' '..'</font> <br />
&#160; <br />
 &#160; <font color="#000099">On Error Resume Next</font> <font color="#006600">' Ignore file errors</font> <br />
 &#160; <font color="#660000">sFile = Dir$(sRoot & "*.*", vbNormal)</font> <br />
 &#160; <font color="#000099">Do While</font> <font color="#660000">LenB(sFile) <br />
 &#160; &#160; &#160;SetAttributes sRoot & sFile, vbNormal <br />
 &#160; &#160; &#160;Kill sRoot & sFile <br />
 &#160; &#160; &#160;sFile = Dir$</font> <br />
 &#160; <font color="#000099">Loop</font> <br />
&#160; <br />
 &#160; On Error GoTo</font> HandleIt <font color="#006600">' No error should occur in here</font> <br />
 &#160; <font color="#000099">Do:</font> <font color="#660000">sDir = Dir$(sRoot & "*", vbDirectory)</font> <br />
 &#160; &#160; &#160;<font color="#000099">For</font> <font color="#660000">iIdx</font> = 1& <font color="#000099">To</font> <font color="#660000">iCnt <br />
 &#160; &#160; &#160; &#160; sDir = Dir$</font> <font color="#006600">'.' '..' ['fail']</font> <br />
 &#160; &#160; &#160;<font color="#000099">Next <br />
 &#160; &#160; &#160;If</font> <font color="#660000">LenB(sDir)</font> = 0& <font color="#000099">Then Exit Do <br />
 &#160; &#160; &#160;If</font> <font color="#660000">KillFolder(sRoot & sDir & DIR_SEP)</font> <font color="#000099">Then <br />
 &#160; &#160; &#160;<font color="#006600">' Sub-folder is now gone but Dir$ was reset <br />
 &#160; &#160; &#160;' during recursive call so Do Dir$(..) again</font> <br />
 &#160; &#160; &#160;Else:</font> <font color="#660000">iCnt = iCnt</font> + 1& <br />
 &#160; &#160; &#160;<font color="#006600">' Kill folder failed (remnant files) so skip <br />
 &#160; &#160; &#160;' this folder (iCnt + 1) to get the rest</font> <br />
 &#160; &#160; &#160;<font color="#000099">End If <br />
 &#160; Loop <br />
&#160; <br />
 &#160; If</font> <font color="#660000">bJustEmptyDontRemove</font> = <font color="#000099">False Then</font> <font color="#660000">RmDir sRoot</font> <font color="#006600">' Errors here if remnants</font> <br />
HandleIt: <br />
 &#160; <font color="#660000">KillFolder</font> = <font color="#000099">Not</font> <font color="#660000">FolderExists(sSpec)</font> <br />
<font color="#000099">End Function <br />
&#160; <br />
&#160;'-----------------------------------------------------<br />
&#160; <br />
&#160; <br />
</font>
</nobr>
</b>
</p>
</tt>

