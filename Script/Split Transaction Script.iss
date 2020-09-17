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
Dim PrimaryDatabaseName As String 

Sub Main
	Call Filename()
	Call CallScriptForPcardStatment()
	Call DatePicker()
	Call FirstDateFilter()	
	Call DirectExtraction1()	
	Call Summarization()	
	Call DirectExtraction2()
	Client.CloseAll
	Call ExportDatabaseXLSX()
	Client.RefreshFileExplorer
End Sub

'
Function Filename()
	'This name will be used for the subDBNames
	Dim MyDate As String 
	MyDate = CStr(Date)
	Dim auditName As String
	auditName = "Split Transaction Audit " + MyDate
	subFilename = InputBox("Type The Name of Audit: ", "Name Input", auditName)
End Function


'This calls a script that will loop through pcard statements and append them together
Function CallScriptForPcardStatment
	Client.RunIDEAScriptEx "Z:\2020 Activities\Data Analytics\Active Scripts\Master Scripts\Loop Pull and Join.iss", "", "", "", ""
	PrimaryDatabaseName = "Append Databases.IMD"
End Function


'
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
			endDate = iRemove(endDate, "/") 'using the @remove date function and replace the slashes whith a blank
		End If 
End Function  

' Data: Direct Extraction. Flitters what is not needed in the first database. Must change date manually. The date is where the main database will start
Function FirstDateFilter()
	Set db = Client.OpenDatabase(PrimaryDatabaseName)
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


' Data: Direct Extraction. Filter what is not needed in the database. Must change date manually. The date is where the second database will start
Function DirectExtraction1
	Set db = Client.OpenDatabase(PrimaryDatabaseName)
	Set task = db.Extraction
	customdbName2 = "Split2-" + subFilename  + ".IMD"
	task.AddExtraction customdbName2, "", "TRANSACTION_DATE >"""  & endDate &  """"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (customdbName2)
End Function


' Analysis: Summarization. Takes all of the transactions with the same vendor and puts them together. 
Function Summarization
	Set db = Client.OpenDatabase(PrimaryDatabaseName)
	Set task = db.Summarization
	task.AddFieldToSummarize "NAME"
	task.AddFieldToSummarize "MERCHANT_NAME"
	task.AddFieldToTotal "TRANSACTION_AMOUNT"
	PrimaryDatabaseName = "Summarization-" + subFilename  + ".IMD"
	task.OutputDBName = PrimaryDatabaseName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (PrimaryDatabaseName)
End Function


' Data: Direct Extraction. Removes everything that is under 5000 and addes a field to indicate if its been worked on or not
Function DirectExtraction2
	Set db = Client.OpenDatabase(PrimaryDatabaseName)
	Set task = db.Extraction
	task.IncludeAllFields
	task.AddField "EDIT", "", WI_EDIT_CHAR, 1, 0, """N"""
	PrimaryDatabaseName = "audit-" + subFilename  + ".IMD"
	task.AddExtraction PrimaryDatabaseName, "", " NO_OF_RECS > 1  .AND.  TRANSACTION_AMOUNT_SUM > 4999"
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (PrimaryDatabaseName)
End Function


' File - Export Database: XLSX. Reorganizes the db and then exports it.
Function ExportDatabaseXLSX
	Set db = Client.OpenDatabase(PrimaryDatabaseName)
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







