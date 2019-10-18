Begin Dialog SplitName 47,103,166,72,"Split Name", .displayIt
  TextBox 58,10,90,16, .NameEnter
  Text 0,10,53,10, "Enter Name of Split", .Text1
  OKButton 5,30,40,14, "OK", .OKButton1
  CancelButton 99,30,40,14, "Cancel", .CancelButton1
End Dialog
Dim filename As String 

Sub Main
	Call ExcelImport()	'C:\Users\mckinnin.lloyd\Documents\Active Projects\P-card split\2019JulyTransactionStatement.xlsx
	Call DirectExtraction()	'2019JulyTransactionStatement-Sheet1.IMD
	Call Summarization()	'Split.IMD
	Call DirectExtraction1()	'Summarization.IMD
	Call ExportDatabaseXLSX()	'Over 4998.IMD
End Sub


' File - Import Assistant: Excel
Function ExcelImport
	Set task = Client.GetImportTask("ImportExcel")
	dbName = "C:\Users\mckinnin.lloyd\Documents\Active Projects\P-card split\2019JulyTransactionStatement.xlsx"
	task.FileToImport = dbName
	task.SheetToImport = "Sheet1"
	task.OutputFilePrefix = "2019JulyTransactionStatement"
	task.FirstRowIsFieldName = "TRUE"
	task.EmptyNumericFieldAsZero = "TRUE"
	task.PerformTask
	dbName = task.OutputFilePath("Sheet1")
	Set task = Nothing
	Client.OpenDatabase(dbName)
End Function

' Data: Direct Extraction
Function DirectExtraction
	Set db = Client.OpenDatabase("2019JulyTransactionStatement-Sheet1.IMD")
	Set task = db.Extraction
	task.AddFieldToInc "NAME"
	task.AddFieldToInc "TRANSACTION_DATE"
	task.AddFieldToInc "TRANSACTION_AMOUNT"
	task.AddFieldToInc "MERCHANT_CATEGORY_CODE_DESCRIPTION"
	task.AddFieldToInc "MERCHANT_NAME"
	filename   = InputBox("Type your name: ", "Name Input", "Split")
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
	task.AddFieldToSummarize "TRANSACTION_DATE"
	task.AddFieldToSummarize "TRANSACTION_AMOUNT"
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
	task.AddExtraction filename + ".IMD", "", "TRANSACTION_AMOUNT_SUM > 4998"
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (filename + ".IMD")
End Function

' File - Export Database: XLSX
Function ExportDatabaseXLSX
	Set db = Client.OpenDatabase(filename + ".IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	task.AddKey "NO_OF_RECS", "D"
	eqn = ""
	task.PerformTask "C:\Users\mckinnin.lloyd\Documents\Active Projects\P-card split\" + filename + ".XLSX", "Database", "XLSX", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function