Begin Dialog DateEntry 50,47,220,150,"Date Entry", .DisplayDateDialog
  Text 16,12,63,14, "Enter Start Date:", .Text1
  Text 16,49,62,16, "Enter End Date:", .Text1
  TextBox 107,11,64,14, .StartDate
  TextBox 107,50,64,14, .EndDate
  OKButton 33,92,40,14, "OK", .OKButton1
  CancelButton 130,92,40,14, "Cancel", .CancelButton1
End Dialog

Begin Dialog ProjectName 50,48,188,121,"Project Name", .DisplayProjName
  Text 12,8,85,10, "Type The Name of Audit", .Text1
  TextBox 12,29,104,14, .NameBox
  OKButton 12,58,40,14, "OK", .OKButton1
  CancelButton 79,58,40,14, "Cancel", .CancelButton1
End Dialog





'Version 1 new script to test for split transactions in pcard purchases
'Version 2 update to improve dialogue boxes
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
Dim dateEntryDialog As DateEntry
Dim button As Integer

Dim errorMessage As String

Sub Main
	Call ProjNameMenu()
	Call ScriptForPcardStatment()
	Call DateChecker()
	Client.CloseAll
	Call StartDateFilter()	
	Call EndDateFilter()	
	Call Summarization()	
	Client.CloseAll
	Call RemoveTransactionsUnder5000()
	Client.RefreshFileExplorer
End Sub

Function ProjNameMenu()
	button = Dialog(projNameDialog)
	If button = 0 Then
		MsgBox("IDEA macro is stopping")
		Stop
	End If		
End Function

Function DateEntryMenu()
		button = Dialog(dateEntryDialog)
		If button = 0 Then
			MsgBox("IDEA macro is stopping")
			Stop
		Else
			Call DateChecker()
		End If
End Function 


Function DisplayProjName(ControlID$, Action%, SuppValue%)
	Dim bExitFunction As Boolean 
	Dim currentDate As String
	'currentDate = CStr(Date())
	Select Case Action%
		Case 1
		
		Case 2
			Select Case ControlID$
				Case "OKButton1"
					If projNameDialog.NameBox = "" Then
						projNameDialog.NameBox = "Split Transaction" 
						subFilename = projNameDialog.NameBox
						MsgBox("Default name is " + subFilename)
					Else 
						subFilename = projNameDialog.NameBox
					End If 
					bExitFunction = True
				Case "CancelButton1"
					bExitFunction = True
			End Select
	End Select 
	
	If bExitFunction Then 
		DisplayProjName = 0
	Else
		DisplayProjName = 1
	End If

End Function

Function DisplayDateDialog(ControlID$, Action%, SuppValue%)
	Dim bExitFunction As Boolean 
	Dim currentDate As String
	'Place holder for text entry boxs
	Select Case Action%
		Case 1
		
		Case 2
			Select Case ControlID$
				Case "OKButton1"
					bExitFunction = True
				Case "CancelButton1"
					bExitFunction = True
			End Select
	End Select 
	
	If bExitFunction Then 
		DisplayDateDialog = 0
	Else
		DisplayDateDialog = 1
	End If

End Function

Function DateChecker() 
	'Place holder for text entry boxs
	dateEntryDialog.StartDate = "YYYY/MM/DD"
	dateEntryDialog.EndDate = "YYYY/MM/DD"
	
	Menu:
		dateError = false
		button = Dialog(dateEntryDialog)
		If button = 0 Then
			MsgBox("This macro has been stopped")
			Stop
		End If
		startDate = dateEntryDialog.StartDate
		endDate = dateEntryDialog.EndDate
	
		'checks to see if the input is in the right formate
		If Not IsDate(startDate) Then
			MsgBox("That is not a correct date in the start date box.")
			dateError = true
		End If
		If Not IsDate(endDate) Then
			MsgBox("That is not a correct date in the end date box.")
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




'This calls a script that will loop through pcard statements and append them together
Function ScriptForPcardStatment

	On Error GoTo ErrorHandler
	'TODO make error check if the file cant be reached. 
	Dim filename As String
	Dim obj As Object
	' Access the CommomDialogs object.
	Set obj = Client.CommonDialogs
	filename = obj.FileOpen("","","All Files (*.*)|*.*||;")
	Client.RunIDEAScriptEx filename, "", "", "", ""
	'Client.RunIDEAScriptEx "Z:\2020 Activities\A.04.2020 - Continuous Audits\Data Analytics\Active Scripts\Master Scripts\Loop Pull and Join.iss", "", "", "", ""
		'TODO fix append error if one already is there
	
	PrimaryDatabaseName = "Append Databases.IMD"
	Set obj = Nothing
	Exit Sub
	ErrorHandler:
		MsgBox "Idea script Loop Pull and Join could not be run. IDEA script stopping."
		Stop
End Function


' Data: Direct Extraction. Flitters what is not needed in the first database. Must change date manually. The date is where the main database will start
Function StartDateFilter
	Dim db As Database
	Dim task As task
	Set db = Client.OpenDatabase(PrimaryDatabaseName)
	Set task = db.Extraction
	task.AddFieldToInc "NAME"
	task.AddFieldToInc "TRANSACTION_DATE"
	task.AddFieldToInc "TRANSACTION_AMOUNT"
	task.AddFieldToInc "MERCHANT_NAME"
	customdbName =  "Split-" + subFilename  + ".IMD"
	task.AddExtraction customdbName, "", "TRANSACTION_DATE >"""  & startDate &  """"
	
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	PrimaryDatabaseName = customdbName
	Client.OpenDatabase (customdbName)
End Function


' Data: Direct Extraction. Filter what is not needed in the database. Must change date manually. The date is where the second database will start
Function EndDateFilter
	Dim db As Database
	Dim task As task

	Set db = Client.OpenDatabase(PrimaryDatabaseName)
	Set task = db.Extraction
	customdbName2 = "Split2-" + projNameDialog.NameBox  + ".IMD"
	task.AddExtraction customdbName2, "", "TRANSACTION_DATE <"""  & endDate &  """"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	PrimaryDatabaseName = customdbName2
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
	PrimaryDatabaseName = "Summarization-" + projNameDialog.NameBox  + ".IMD"
	task.OutputDBName = PrimaryDatabaseName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (PrimaryDatabaseName)
End Function


' Data: Direct Extraction. Removes everything that is under 5000 and addes a field to indicate if its been worked on or not
Function RemoveTransactionsUnder5000
	Dim db As Database
	Dim task As task
	Set db = Client.OpenDatabase(PrimaryDatabaseName)
	Set task = db.Extraction
	task.IncludeAllFields
	task.AddField "EDIT", "", WI_EDIT_CHAR, 1, 0, """N"""
	PrimaryDatabaseName = "audit-" + projNameDialog.NameBox  + ".IMD"
	task.AddExtraction PrimaryDatabaseName, "", " NO_OF_RECS > 1  .AND.  TRANSACTION_AMOUNT_SUM > 4999"
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (PrimaryDatabaseName)
End Function










