Sub XMLDriver()
	on error resume next

	Dim XMLDoc
	Dim ErrorMsg
	Dim ProjectFile
	Dim RunGroupName

	If WScript.Arguments.Count>=2 then
		ProjectFile=Trim(WScript.Arguments(0))
		RunGroupName=Trim(WScript.Arguments(1))
		
		RunGroupNameArray = Split(RunGroupName, ",")
	Else
		wscript.Quit 400
	End If

	set XMLDoc=wscript.CreateObject("MSXML2.DOMDOCUMENT")

	XMLDoc.async=False

	XMLDoc.ValidateonParse=True
	'Open project file
	Call XMLDoc.load(ProjectFile)

	If XMLDoc.parseError.errorCode <> 0 Then
		ErrorMsg = "Reason:" + Chr(9) + XMLDoc.parseError.reason + Chr(13) + Chr(10) + _
			"Line:" + Chr(9) + CStr(XMLDoc.parseError.line) + Chr(13) + Chr(10) + _
			"Pos:" + Chr(9) + CStr(XMLDoc.parseError.linePos) + Chr(13) + Chr(10) + _
			"Source:" + Chr(9) + XMLDoc.parseError.srcText
		' Post an error to the log and exit
		Wscript.echo("Cannot parse the document:" + ErrorMsg) 
		wscript.quit 500
	End If

	Set RootNode=XMLDoc.documentElement
	
	'Find a particular element using XPath:
	Set GroupNode=XMLDOC.selectSingleNode("//Nodes/Node/Node[@name='test data']")
	
	'loop to find the specified Node according to the Group Name
	for i = 0 to Ubound(RunGroupNameArray)
	
		Call ProcessNode(GroupNode,RunGroupNameArray(i))
		
	Next
	'Close and save the project file
	XMLDOC.save(ProjectFile)

End Sub

'===============================================================================
Sub ProcessNode(aNode,aGroupName)
	for i = 0 to aNode.childnodes.length - 1
		set childlist=anode.childnodes(i)
		for j = 0 to childlist.attributes.length - 1
			if UCase(childlist.attributes.item(j).nodename) = UCase("value") and UCase(Left(childlist.attributes.item(j).nodevalue,2)) = UCase(Left(Trim(aGroupName),2))then
				'set attribute enabled=-1
				set bothernode=anode.childnodes(i-4)
				call bothernode.setAttribute("value","-1")
			end if
		next
		Call ProcessNode(childlist,aGroupName)
	Next
	
End Sub

'================================================================================
call XMLDriver()