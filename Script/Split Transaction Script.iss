Begin Dialog DateEntry 50,49,221,150,"Date Entry", .DisplayDateDialog
  Text 36,12,63,14, "Enter Start Date:", .Text1
  Text 17,12,63,14, "Enter Start Date:", .Text1
  Text 17,50,78,15, "Enter End Date:", .Text1
  TextBox 107,12,64,14, .TextBox1
  TextBox 107,50,64,14, .TextBox2
  OKButton 33,94,40,14, "OK", .OKButton1
  CancelButton 129,94,40,14, "Cancel", .CancelButton1
End Dialog

Begin Dialog ProjectName 50,48,188,121,"Project Name", .DisplayProjName
  Text 12,8,85,10, "Type The Name of Audit", .Text1
  TextBox 12,29,104,14, .NameBox
  OKButton 12,58,40,14, "OK", .OKButton1
  CancelButton 79,58,40,14, "Cancel", .CancelButton1
End Dialog




Option Explicit

Dim customdbName As String 
Dim customdbName2 As String 
Dim dbName As String
Dim subFilename As String
Dim startDate  As String
Dim endDate As String
Dim dateError As Boolean
Dim PrimaryDatabaseName As String 
Dim bExitScript As Boolean 

Dim projNameDialog As ProjectName


Sub Main
	Call mainMenu()
	If Not bExitScript Then 
		'Call Filename()
		'Call CallScriptForPcardStatment()
		'Call DatePicker()
		'Call FirstDateFilter()	
		'Call DirectExtraction1()	
		'Call Summarization()	
		'Call DirectExtraction2()
		'Client.CloseAll
		'Call ExportDatabaseXLSX()
		'Client.RefreshFileExplorer
		MsgBox("The script ending")
	Else 
		MsgBox("The script has been canculed")
	End If
End Sub

Function mainMenu()
	Dim button As Integer 
	
	button = Dialog(projNameDialog)
End Function


Function DisplayProjName(ControlID$, Action%, SuppValue%)
	Dim bExitFunction As Boolean 
	Dim currentDate As String
	currentDate = CStr(Date())
	Select Case Action%
		Case 1
		
		Case 2
			Select Case ControlID$
				Case "OKButton1"
					If projNameDialog.NameBox = "" Then
						projNameDialog.NameBox = "Split Transaction " + currentDate
						MsgBox("Default name is " + projNameDialog.NameBox)
					End If 
					bExitFunction = True
				Case "CancelButton1"
					bExitScript = True
					bExitFunction = True
			End Select
	End Select 
	
	If bExitFunction Then 
		DisplayProjName = 0
	Else
		DisplayProjName = 1
	End If

End Function
'
Function Filename()
	'This name will be used for the subDBNames
	subFilename = InputBox("Type The Name of Audit: ", "Name Input", "Split Transaction")
	MsgBox(subFilename)
	
End Function


'This calls a script that will loop through pcard statements and append them together
Function CallScriptForPcardStatment
	Client.RunIDEAScriptEx "Z:\2020 Activities\Data Analytics\Active Scripts\Master Scripts\Loop Pull and Join.iss", "", "", "", ""
	PrimaryDatabaseName = "Append Databases.IMD"
End Function


'
Function DatePicker()
	Dim dateDialog As DateEntry
	Dim button As Integer
	'Place holder for text entry boxs
	dateEntry.TextBox1 = "YYYY/MM/DD"
	dateEntry.TextBox2 = "YYYY/MM/DD"
	
	Menu:
		dateError = false
		button = Dialog(dateEntry)
		startDate = dateEntry.TextBox1
		endDate = dateEntry.TextBox2
	
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
	Dim db As Database
	Dim task as task
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
	Dim db As Database
	Dim task As task
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
	Dim db As Database
	Dim task As task
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
	Dim db As Database
	Dim task As task
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
	Dim db As Database
	Dim task as task
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









