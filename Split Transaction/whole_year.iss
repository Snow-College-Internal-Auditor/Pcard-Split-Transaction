Begin Dialog DateEntry 50,49,221,150,"Date Entry", .NewDialog
  Text 36,12,63,14, "Enter Start Date:", .Text1
  Text 17,12,63,14, "Enter Start Date:", .Text1
  Text 17,50,78,15, "Enter End Date:", .Text1
  TextBox 107,12,64,14, .TextBox1
  TextBox 107,50,64,14, .TextBox2
  OKButton 33,94,40,14, "OK", .OKButton1
  CancelButton 129,94,40,14, "Cancel", .CancelButton1
End Dialog
Dim customdbName As String 
Dim customdbName2 As String 
Dim dbName As String
Dim subFilename As String
Dim dlg As DateEntry
Dim button As Integer
Dim startDate  As String
Dim endDate As String
Dim dateError As Boolean

Sub Main
	Call Filename()
	Call ExcelImport()
	Call DatePicker()
	Call FirstDateFilter()	
	Call ExcelImport1()	
	Call DirectExtraction1()	
	Call AppendDatabase()	
	Call Summarization()	
	Call DirectExtraction2()
	Client.CloseAll
	Call ExportDatabaseXLSX()
	Client.RefreshFileExplorer
End Sub


Function Filename()
	'This name will be used for the subDBNames
	subFilename = InputBox("Type The Name of The Month: ", "Name Input", "Month")
End Function

' File - Import Assistant: Excel
Function ExcelImport
	Set task = Client.GetImportTask("ImportExcel")
	Set obj = client.commondialogs
		dbName =  obj.fileopen("","","All Files (*.*)|*.*||;")
	task.FileToImport = dbName
	dbName = dbName + subFilename
	task.SheetToImport = "Sheet1"
	task.OutputFilePrefix = iSplit(dbName ,"","\",1,1)
	task.FirstRowIsFieldName = "TRUE"
	task.EmptyNumericFieldAsZero = "TRUE"
	task.PerformTask
	dbName = task.OutputFilePath("Sheet1")
	Set task = Nothing
	Client.RefreshFileExplorer
End Function

Function DatePicker() 
	'Place holder for text entry boxs
	dlg.TextBox1 = "YYYY/MM/DD"
	dlg.TextBox2 = "YYYY/MM/DD"
	
	Menu:
		dateError = false
		button = Dialog(dlg)
		startDate = dlg.TextBox1
		endDate = dlg.TextBox2
	
		'checks to see if the input is in the right formate
		If Not IsDate(startDate) Then
			MsgBox("That is not a correct date")
			dateError = true
		End If
		If Not IsDate(endDate) Then
			MsgBox("That is not a correct date")
			dateError = true
		End If
		'if it is not in the right formate it will loop back to Menu:
		If dateError = true Then 
			GoTo Menu
		Else
			'remove the slashes from the dates as IDEA date does not use them
			startDate = iRemove(startDate, "/") 'using the @remove date function and replace the slashes whith a blank
			MsgBox(startDate)
			endDate = iRemove(endDate, "/") 'using the @remove date function and replace the slashes whith a blank
			MsgBox(endDate)
		End If 
		
		'If button = 2 Then 
		'	bExitScript = True 
		'ElseIf button = 1 Then 
		'	MsgBox(startDate)
		'End If 
	
End Function  

' Data: Direct Extraction. Flitters what is not needed in the first database. Must change date manually. The date is where the main database will start
Function FirstDateFilter()
	Set db = Client.OpenDatabase(dbName)
	Set task = db.Extraction
	task.AddFieldToInc "NAME"
	task.AddFieldToInc "TRANSACTION_DATE"
	task.AddFieldToInc "TRANSACTION_AMOUNT"
	task.AddFieldToInc "MERCHANT_NAME"
	customdbName =  "Split-" + subFilename + ".IMD"
	task.AddExtraction customdbName, "", "TRANSACTION_DATE >"""  & startDate &  """"

	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (customdbName)
End Function

' File - Import Assistant: Excel Brings in the second database tat will join with the first database
Function ExcelImport1
	Set task = Client.GetImportTask("ImportExcel")
	Set obj = client.commondialogs
		dbName =  obj.fileopen("","","All Files (*.*)|*.*||;")
	task.FileToImport = dbName
	task.SheetToImport = "Sheet1"
	task.OutputFilePrefix = iSplit(dbName ,"","\",1,1)
	task.FirstRowIsFieldName = "TRUE"
	task.EmptyNumericFieldAsZero = "TRUE"
	task.PerformTask
	dbName = task.OutputFilePath("Sheet1")
	Set task = Nothing
	Client.RefreshFileExplorer
End Function


' Data: Direct Extraction. Filter what is not needed in the database. Must change date manually. The date is where the second database will start
Function DirectExtraction1
	Set db = Client.OpenDatabase(dbName)
	Set task = db.Extraction
	task.AddFieldToInc "NAME"
	task.AddFieldToInc "TRANSACTION_DATE"
	task.AddFieldToInc "TRANSACTION_AMOUNT"
	task.AddFieldToInc "MERCHANT_NAME"
	customdbName2 = "Split2-" + subFilename  + ".IMD"
	task.AddExtraction customdbName2, "", "TRANSACTION_DATE >"""  & endDate &  """"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (customdbName2)
End Function

' File: Append Databases. Appends the split2 and splt db together to be fillttered
Function AppendDatabase
	Set db = Client.OpenDatabase(customdbName)
	Set task = db.AppendDatabase
	task.AddDatabase customdbName2
	dbName = "Append Databases-" + subFilename + ".IMD"
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Analysis: Summarization. Takes all of the transactions with the same vendor and puts them together. 
Function Summarization
	Set db = Client.OpenDatabase(dbName)
	Set task = db.Summarization
	task.AddFieldToSummarize "NAME"
	task.AddFieldToSummarize "MERCHANT_NAME"
	task.AddFieldToTotal "TRANSACTION_AMOUNT"
	dbName = "Summarization-" + subFilename  + ".IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Data: Direct Extraction. Removes everything that is under 5000 and addes a field to indicate if its been worked on or not
Function DirectExtraction2
	Set db = Client.OpenDatabase(dbName)
	Set task = db.Extraction
	task.IncludeAllFields
	task.AddField "EDIT", "", WI_EDIT_CHAR, 1, 0, """N"""
	dbName = "audit-" + subFilename  + ".IMD"
	task.AddExtraction dbName, "", " NO_OF_RECS > 1  .AND.  TRANSACTION_AMOUNT_SUM > 4999"
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' File - Export Database: XLSX. Reorganizes the db and then exports it.
Function ExportDatabaseXLSX
	Set db = Client.OpenDatabase(dbName)
	Set task = db.Index
	task.AddKey "NO_OF_RECS", "D"
	task.Index FALSE
	task = db.ExportDatabase
	task.IncludeAllFields
	' Display the setup dialog box before performing the task.
	task.DisplaySetupDialog 0
	Set db = Nothing
	Set task = Nothing
End Function





