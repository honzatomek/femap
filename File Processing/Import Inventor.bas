'#Reference {780FB82C-A4AA-4043-A0B6-ABCA50DB747A}#1.0#0#C:\Program Files\Autodesk\Inventor 2011\Bin\RxApprentice.tlb#Autodesk Inventor's Apprentice Object Library#InventorApprentice
' This VB script will convert a Autocad Inventor file (part or assembly) to ACIS SAT
'   and import the geometry into Femap
' This script requires Inventor viewer to be installed and 
'   Inventor's Apprentice Object Library is properly referenced.

Sub Main
    Dim App As femap.model
    Set App = feFemap()
	Dim iFile As String
	Dim oFile As String

	Dim iApp As InventorApprentice.ApprenticeServerComponent
	Dim cDef As InventorApprentice.ComponentDefinition
	Dim oDoc As InventorApprentice.ApprenticeServerDocument
	Dim dio As InventorApprentice.DataIO

    Dim formatOK As Boolean
	Dim ts() As String
	Dim store() As StorageTypeEnum

	Set iApp = Nothing

	Set iApp = New InventorApprentice.ApprenticeServerComponent
	If( iApp Is Nothing ) Then
		App.feAppMessage(FCM_ERROR, "Required Inventor viewer not found" )
		Exit All
	End If
	'Dim vs As String
	'vs = iApp.SoftwareVersion().DisplayVersion

	App.feFileGetName("Read File", "Choose Inventor File", "*.ipt;*.iam",True, iFile)
	Set oDoc=iApp.Open(iFile)
	If( IsNull(oDoc) ) Then
		App.feAppMessage(FCM_ERROR, "Unable to open document: " + iFile)
		iApp.Close()
		Exit All
	End If

	Set cDef = oDoc.ComponentDefinition
	Set dio = cDef.DataIO

    oFile = Left$(iFile, Len(iFile)-3)+"sat"
    formatOK = False
	dio.GetOutputFormats(ts, store)
	For i = 0 To 8
		If( StrComp("ACIS SAT", ts(i)) And (store(i) = kFileStorage Or store(i) = kFileOrStreamStorage) )Then
			formatOK = True
			Exit For
		End If
	Next i
	If( formatOK) Then
		dio.WriteDataToFile("ACIS SAT", oFile)
		App.feFileReadAcis(oFile)
	Else
		App.feAppMessage(FCM_ERROR, "File: "+iFile+" cannot be translated")
	End If

	iApp.Close()
End Sub
