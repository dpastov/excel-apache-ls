# excel-apache-ls
Create Excel file using LotusScript without MS Excel (OLE) installed on PC

#Example of agent

```
Option Public
Option Declare

UseLSX "*javacon"
Use "Apache.Excel"

Sub Initialize
	Dim jSession As JavaSession
	Dim jClass As Javaclass
	Dim jObject As JavaObject
	Dim filepath As String
	Dim row As Integer

	Set jSession = New Javasession
	Set jClass = jSession.GetClass("explicants.office.Excel")
	Set jObject = jClass.Createobject()
	
	Call jObject.createSheet("sheet A-100")
	Call jObject.createSheet("sheet B-100")
	Call jObject.createSheet("sheet C-100")
	
	Call jObject.getSheet("sheet A-100")

	row = row + 1
	Call jObject.setCellValueString("lorem", row, 0)
	Call jObject.setCellValueString("ipsum", row, 1)
	Call jObject.setCellValueDouble(55, row, 2)
	
	row = row + 1
	Call jObject.setCellValueString("hello", row, 0)
	Call jObject.setCellValueString("world", row, 1)
	Call jObject.setCellValueDouble(200.50, row, 2)
	
	row = row + 1
	Call jObject.setCellValueString("gurli gris", row, 0)
	Call jObject.setCellValueString("george", row, 1)
	Call jObject.setCellValueDouble(0.505, row, 2)
	
	filepath = temp() & Join(Evaluate({@Unique})) & ".xls"
	Call jObject.saveAsFile(filepath)
	
	MsgBox filepath
End Sub
Sub Terminate
	
End Sub


Private Function temp() As String
	Dim tmpDir As String
	tmpdir = Environ("TEMP")
	If tmpdir = "" Then
		tmpdir = Environ("TMP")
	End If
	
	If Right$(tmpdir, 1) <> "\" Then
		tmpdir = tmpdir & "\"
	End If
	
	temp = tmpdir
End Function
```
