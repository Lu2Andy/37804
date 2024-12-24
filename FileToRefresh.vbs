ConfigFilePath = "C:\VBS\AutoRefreshConfig.xlsx"

Set objShell = CreateObject("Wscript.Shell")
'strPath = Wscript.ScriptFullName 
strPath = Wscript.ScriptFullName 
Set objShell = Nothing
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(strPath) 
strFileName = objFSO.GetFileName(objFile) 
Dim pos
pos = InStrRev(strFileName, ".")
TaskName = Left(strFileName, pos - 1)



Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(ConfigFilePath, False, True)
Set ojbWorksheet = objWorkbook.WorkSheets("Config")
LogFile = ojbWorksheet.Cells(2, 17)
Set objLog = objFSO.OpenTextFile(LogFile, 8, 1, -1)
objLog.WriteLine Now & "," & TaskName & " Program start, ,"
FoundTask = False
RefreshError = False

RowCount = 2
For RowCount = 2 To ojbWorksheet.Range(ojbWorksheet.Cells(1, 1), ojbWorksheet.UsedRange).Rows.Count
	if LCase(Trim(ojbWorksheet.Cells(RowCount, 2))) = LCase(Trim(TaskName)) Then
		FoundTask = True
		VarCategroy = ojbWorksheet.Cells(RowCount, 3)
		VarSourcePath = ojbWorksheet.Cells(RowCount, 4)
		VarDestinationPath = ojbWorksheet.Cells(RowCount, 5)
		VarOwner = ojbWorksheet.Cells(RowCount, 7)
		' VarDtCheckMon = ojbWorksheet.Cells(RowCount, 8)
		' VarDtCheckTue = ojbWorksheet.Cells(RowCount, 9)
		' VarDtCheckWed = ojbWorksheet.Cells(RowCount, 10)
		' VarDtCheckThu = ojbWorksheet.Cells(RowCount, 11)
		' VarDtCheckFri = ojbWorksheet.Cells(RowCount, 12)
		objLog.WriteLine Now & ",Category=" & VarCategroy & "||SourcePath=" & VarSourcePath & "||DestinationPath=" & VarDestinationPath & ", ,"
	
		exit For
	end if
Next

set ojbWorksheet = Nothing
objWorkbook.Close

if FoundTask = True Then
	objLog.WriteLine Now & "," & TaskName & "," & VarSourcePath & ",start"
end if	

objExcel.Visible = False
objExcel.DisplayAlerts = False
objExcel.AskToUpdateLinks = False
objExcel.AlertBeforeOverwriting = False
On Error Resume Next
If LCase(Trim(VarCategroy)) = "folder" Then
	For Each f In objFSO.GetFolder(VarSourcePath).Files
	  If LCase(objFSO.GetExtensionName(f.Name)) = "xlsx" or LCase(objFSO.GetExtensionName(f.Name)) = "xlsm" Then
		Set objWorkbook = objExcel.Workbooks.Open(f.Path)		
		objLog.WriteLine Now & "," & TaskName & " File <" & f.Name & "> Connections:" & objWorkbook.Connections.Count & ", ,"		
		For Each Cn In objWorkbook.Connections			
			Err.Clear
			bg = Cn.OLEDBConnection.BackgroundQuery
			if Err.Description = "" Then
				Cn.OLEDBConnection.BackgroundQuery = False
			end if
			Cn.Refresh
			if Err.Description = "" Then
				Cn.OLEDBConnection.BackgroundQuery = bg
			Else
				objLog.WriteLine Now & "," & TaskName & " File <" & f.Name & "> soft remind on refresh error:" & Err.Description & ", ,"		
				Err.Clear
			end if			
			WScript.Sleep 1000			
		Next
		WScript.Sleep 1000
		if LCase(Trim(VarSourcePath)) = LCase(Trim(VarDestinationPath)) Then
			objWorkbook.Save
			objLog.WriteLine Now & "," & TaskName & " File <" & f.Name & "> save success, ,"
		Else			
			If objFSO.FolderExists(VarDestinationPath) = True Then
				objWorkbook.SaveAs (objFSO.GetFolder(VarDestinationPath).Path & "\" & f.Name)
				objLog.WriteLine Now & "," & TaskName & " File <" & f.Name & "> SaveAs success, ,"
			Else
				objLog.WriteLine Now & "," & TaskName & " Folder <" & VarDestinationPath & "> does not exist, ,"
				RefreshError = True
			end if
		end if
		objWorkbook.Close
		set objWorkbook = Nothing
	  End If
	Next
Elseif LCase(Trim(VarCategroy)) = "file" Then
	If LCase(objFSO.GetExtensionName(VarSourcePath)) = "xlsx" or LCase(objFSO.GetExtensionName(VarSourcePath)) = "xlsm" Then
		If objFSO.FileExists(VarSourcePath) = True Then
			Set objWorkbook = objExcel.Workbooks.Open(VarSourcePath, False, False)			
			objLog.WriteLine Now & "," & TaskName & " File <" & f.Name & "> Connections:" & objWorkbook.Connections.Count & ", ,"
			For Each Cn In objWorkbook.Connections				
				Err.Clear
				bg = Cn.OLEDBConnection.BackgroundQuery
				if Err.Description = "" Then
					Cn.OLEDBConnection.BackgroundQuery = False
				end if
				Cn.Refresh
				if Err.Description = "" Then
					Cn.OLEDBConnection.BackgroundQuery = bg
				Else
					objLog.WriteLine Now & "," & TaskName & " File <" & f.Name & "> Refresh error:" & Err.Description & ", ,"		
					Err.Clear
				end if
			Next
			WScript.Sleep 1000
			if LCase(Trim(VarSourcePath)) = LCase(Trim(VarDestinationPath)) Then
				objWorkbook.Save
				objLog.WriteLine Now & "," & TaskName & " File save success, ,"
			Else
				objWorkbook.SaveAs (VarDestinationPath)
				objLog.WriteLine Now & "," & TaskName & " File SaveAs success, ,"
			end if
			objWorkbook.Close
			set objWorkbook = Nothing
		Else
			objLog.WriteLine Now & "," & TaskName & " cannot find the file:" & VarSourcePath &", ,"
			RefreshError = True
		end if
	end if
end if
objExcel.Quit
set objExcel = Nothing
Set objFile = Nothing
Set objFSO = Nothing

if Err.Description <> "" Then
	objLog.WriteLine Now & "," & TaskName & " running error: " & Err.Description & ", ,"
Else
	if FoundTask = True and RefreshError = False Then
		objLog.WriteLine Now & "," & TaskName & "," & VarSourcePath & ",end"
	end if	
end if
objLog.WriteLine Now & "," & TaskName & " Program end, ,"

'MsgBox "Done"
