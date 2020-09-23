<div align="center">

## Log \(class\)


</div>

### Description

Easily log events. Log errors by passing the ERR object.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Michael A\. Schmidt](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/michael-a-schmidt.md)
**Level**          |Advanced
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/michael-a-schmidt-log-class__1-25570/archive/master.zip)





### Source Code

```
Option Explicit
'================================
' Michael Schmidt July 2001
' mikes@mtdmarketing.com
'================================
'================================
' Example:
' Public MyLog As Log
'
' Private Form_Load
' On Error Goto ErrorSub
'
' MyLog = New Log
' Log ("Loading Form...")
' Log ("Unloading Form...","Hello!")
'
' Exit Sub
' ErrorSub:
'
' LogError(Err,"Error in MySub")
'
' End Sub
'=================================
' The EVENT function was never
' implemented, if you compile
' this into a DLL then you should
' be able to use the EVENT feature
' quite handy.
'==================================
Private LogFile As Long
Private LogName As String
Private Const Comma = ","
Private Const Quote = """"
Private Const Space = " "
Private oDateTime
Private oType
Private oGeneralInfo
Private oDetailedInfo
Event LogIn(logData As String)
Private Sub LogError(objError As ErrObject, strSubFailed As String)
 oDateTime = "(" & Date & Space & Time & ")"
 oType = "ERROR"
 oGeneralInfo = "Error " & objError.Number & " - " & Err.Description
 oDetailedInfo = strSubFailed
 AppendLog
End Sub
Private Sub Log(strGeneral As String, Optional strDetailed As String)
 oDateTime = "(" & Date & Space & Time & ")"
 oType = "GENERAL"
 oGeneralInfo = strGeneral
 oDetailedInfo = strDetailed
 AppendLog
End Sub
Private Sub AppendLog()
Dim CSVstring As String
Dim BASstring As String
 CSVstring = Quote & oDateTime & Quote & Comma & _
 Quote & oType & Quote & Comma & _
 Quote & oGeneralInfo & Quote & Comma & _
 Quote & oDetailedInfo & Quote
 BASstring = oDateTime & Space & _
 oType & Space & _
 oGeneralInfo & _
 oDetailedInfo
 RaiseEvent LogIn(BASstring)
 ' Print to LOG
 Open LogName For Append As #LogFile
 Print #LogFile, CSVstring
 Close #LogFile
End Sub
Private Sub Class_Initialize()
 LogName = App.Path & "\Session.log"
 LogFile = FreeFile()
 Open LogName For Output As #LogFile
 Close #LogFile
 Log ("[Log Started]")
End Sub
'=================================
' Path Property
'=================================
Property Get LogFilePathName() As String
 LogFilePathName = LogName
End Property
Private Sub Class_Terminate()
 Log ("[Log Ended]")
End Sub
```

