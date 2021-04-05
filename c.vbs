Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("C:\kumar\tst\stateMcCreator\StateMcCreator.xlsm")

' objExcel.Application.Visible = True

Set objArgs = WScript.Arguments
'For I = 0 to objArgs.Count - 1
'   WScript.Echo objArgs(I)
'Next



if objArgs.Count <> 1 then
    WScript.Echo "Usage: " & WScript.ScriptFullName & " <Class Name>"
    WScript.Quit
End if

prog = "StateMcCreator.xlsm!Sheet1.GenCode_Script(" & Chr(34) & objArgs(0) & Chr(34) & ")"
'WScript.Echo prog

objExcel.Application.Run prog
objExcel.ActiveWorkbook.Close
WScript.Quit 