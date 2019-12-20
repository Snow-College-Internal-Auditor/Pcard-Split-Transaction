Dim filename As String 
Dim dbName As String 

Sub Main
	Call ExcelImport()	
	Call DirectExtraction()	
	Call Summarization()	
	Call DirectExtraction1()	
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
	task.AddFieldToInc "MERCHANT_NAME"
	task.AddFieldToInc "TRANSACTION_DATE"
	task.AddFieldToInc "TRANSACTION_AMOUNT"
	filename =  InputBox("Type your name: ", "Name Input", "Split")
	task.AddExtraction filename + ".IMD", "", ""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Client.OpenDatabase ((filename + ".IMD"))
End Function

' Analysis: Summarization
Function Summarization
	Set db = Client.OpenDatabase(filename + ".IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "NAME"
	task.AddFieldToSummarize "MERCHANT_NAME"
	task.AddFieldToTotal "TRANSACTION_AMOUNT"
	filename =  InputBox("Type your name: ", "Name Input", "Summarization1")
	task.OutputDBName = filename + ".IMD"
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (filename + ".IMD")
End Function

' Data: Direct Extraction
Function DirectExtraction1
	Set db = Client.OpenDatabase(filename + ".IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	filename =  InputBox("Type your name: ", "Name Input","Over 4998_1")
	task.AddExtraction filename + ".IMD", "", "NO_OF_RECS > 1  .AND.  TRANSACTION_AMOUNT_SUM > 4998"
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End Function

' File - Export Database: XLSX
Function ExportDatabaseXLSX
	Set db = Client.OpenDatabase(filename + ".IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	task.AddKey "NO_OF_RECS", "D"
	eqn = ""	
	task.DisplaySetupDialog 0
	Set db = Nothing
	Set task = Nothing
End Function


