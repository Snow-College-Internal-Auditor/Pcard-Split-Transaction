Dim customdbName As String 
Dim customdbName2 As String 
Dim dbName As String 

Sub Main
	Call ExcelImport()
	Call DirectExtraction()	
	Call ExcelImport1()	
	Call DirectExtraction1()	
	Call AppendDatabase()	
	Call Summarization()	
	Call DirectExtraction2()	
	Call ExportDatabaseXLSX()
End Sub


' File - Import Assistant: Excel
Function ExcelImport
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
End Function

' Data: Direct Extraction
Function DirectExtraction
	Set db = Client.OpenDatabase(dbName)
	Set task = db.Extraction
	task.AddFieldToInc "NAME"
	task.AddFieldToInc "TRANSACTION_DATE"
	task.AddFieldToInc "TRANSACTION_AMOUNT"
	task.AddFieldToInc "MERCHANT_NAME"
	customdbName =  InputBox("Type your name: ", "Name Input", "Split")
	customdbName =  customdbName + ".IMD"
	task.AddExtraction customdbName,"", "TRANSACTION_DATE  > ""20180630"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (customdbName)
End Function

' File - Import Assistant: Excel
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
End Function


' Data: Direct Extraction
Function DirectExtraction1
	Set db = Client.OpenDatabase(dbName)
	Set task = db.Extraction
	task.AddFieldToInc "NAME"
	task.AddFieldToInc "TRANSACTION_DATE"
	task.AddFieldToInc "TRANSACTION_AMOUNT"
	task.AddFieldToInc "MERCHANT_NAME"
	customdbName2 =  InputBox("Type your name: ", "Name Input", "Split2")
	customdbName2 = customdbName2 + ".IMD"
	task.AddExtraction customdbName2, "", "TRANSACTION_DATE > ""20190630"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End Function

' File: Append Databases
Function AppendDatabase
	Set db = Client.OpenDatabase(customdbName)
	Set task = db.AppendDatabase
	task.AddDatabase customdbName2
	dbName = InputBox("Type your name: ", "Name Input", "Append Databases")
	task.PerformTask dbName + ".IMD", ""
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName + ".IMD")
End Function

' Analysis: Summarization
Function Summarization
	Set db = Client.OpenDatabase(dbName + ".IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "NAME"
	task.AddFieldToSummarize "MERCHANT_NAME"
	task.AddFieldToTotal "TRANSACTION_AMOUNT"
	dbName = InputBox("Type your name: ", "Name Input", "Summarization")
	dbName = dbName + ".IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Data: Direct Extraction
Function DirectExtraction2
	Set db = Client.OpenDatabase(dbName)
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = InputBox("Type your name: ", "Name Input", "audit")
	dbName = dbName + ".IMD"
	task.AddExtraction dbName, "", " NO_OF_RECS > 1  .AND.  TRANSACTION_AMOUNT_SUM > 4998"
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' File - Export Database: XLSX
Function ExportDatabaseXLSX
	Set db = Client.OpenDatabase(dbName)
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask "C:\Users\mckinnin.lloyd\Documents\Active Projects\P-card split\User input\" + dbName + ".xlsx", "Database", "XLSX", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function