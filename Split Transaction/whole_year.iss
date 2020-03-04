Dim customdbName As String 
Dim customdbName2 As String 
Dim dbName As String
Dim subFilename As String

Sub Main
	Call Filename()
	Call ExcelImport()
	Call DirectExtraction()	
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

' Data: Direct Extraction. Flitters what is not needed in the first database. Must change date manually. The date is where the main database will start
Function DirectExtraction
	Set db = Client.OpenDatabase(dbName)
	Set task = db.Extraction
	task.AddFieldToInc "NAME"
	task.AddFieldToInc "TRANSACTION_DATE"
	task.AddFieldToInc "TRANSACTION_AMOUNT"
	task.AddFieldToInc "MERCHANT_NAME"
	customdbName =  "Split-" + subFilename + ".IMD"
	task.AddExtraction customdbName, "", "TRANSACTION_DATE > ""20180731"""
	task.CreateVirtualDatabase = False
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
	task.AddExtraction customdbName2, "", "TRANSACTION_DATE > ""20190731"""
	task.CreateVirtualDatabase = False
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





