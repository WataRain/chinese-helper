Attribute VB_Name = "ZhuyinMacro"
Sub GuoyinPlease()
'
' GuoyinPlease Macro
'
'

Dim pythonExe, pythonScript As String

'For consideration: A menu to set these?
'!-Change these strings to match your install dir------------------------------------!
pythonExe = """C:\Python\chinese-helper\Scripts\python.exe"""
guoyinPy = """C:\Python\chinese-helper\guoyin.py"""
'!-----------------------------------------------------------------------------------!

documentPath = """" & ActiveDocument.FullName & """"

' Execute guoyin.py, read standard output/error streams
Dim shellCommand As String
    shellCommand = pythonExe & " " & guoyinPy & " " & documentPath
Dim objShell As Object
    Set objShell = VBA.CreateObject("Wscript.Shell")
    
Debug.Print (shellCommand)
    Set objShellExec = objShell.Exec(shellCommand)

'shellRuntime = 0
'While objShellExec.Status = 0
'    shellRuntime = shellRuntime + 1
'    If shellRuntime Mod 100 = 0 Then
'        DoEvents
'    End If
'Wend

' Apply Zhuyin as phonetic guide
Dim currentRange As Range
For wordCount = 1 To ActiveDocument.Characters.Count - 1
    zhuyin = objShellExec.StdOut.ReadLine
    Set currentRange = ActiveDocument.Characters(wordCount)
    
    If zhuyin <> "x" Then
        With currentRange
            .PhoneticGuide Text:=zhuyin, FontSize:=.Font.Size * 0.5, FontName:=.Font.Name, Alignment:=wdPhoneticGuideAlignmentRightVertical
        End With
        Debug.Print (wordCount & " " & currentRange & " " & zhuyin)
    End If
    
    If wordCount Mod 10 = 0 Then
        DoEvents
    End If
    
Next wordCount

End Sub
