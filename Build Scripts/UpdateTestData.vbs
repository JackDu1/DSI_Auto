'==============================DSI_FinishInstall_ToadforOracle========================================
Function Update_DSI_FinishInstall_ToadforOracle(StrProduct,StrVersion)

	Dim Conn,isSQL
	on error resume next
	
	if IsEmpty(StrProduct) then
		Update_DSI_FinishInstall_ToadforOracle=false
		wscript.quit 100
	else
		select case StrProduct
			case "TOADFORORACLE_X64_EN"
				StrProduct="64-bit"
			case "TOADFORORACLE_X64_ZH"
				StrProduct="64-bit"
			case "TOADFORORACLE_X86_EN"
				StrProduct="32-bit"
			case "TOADFORORACLE_X86_ZH"
				StrProduct="32-bit"
			case "TOADFORORACLE_TRIAL_X86_EN"
				StrProduct="32-bit Trial"
			case "TOADFORORACLE_TRIAL_X86_ZH"
				StrProduct="32-bit Trial"
			case "TOADFORORACLE_TRIAL_X64_EN"
				StrProduct="64-bit Trial"
			case "TOADFORORACLE_TRIAL_X64_ZH"
				StrProduct="64-bit Trial"
			case "TOADFORORACLE_READONLY_X86_EN"
				StrProduct="32-bit Read-Only"
			case "TOADFORORACLE_READONLY_X86_ZH"
				StrProduct="32-bit Read-Only"
			case "TOADFORORACLE_READONLY_X64_EN"
				StrProduct="64-bit Read-Only"
			case "TOADFORORACLE_READONLY_X64_ZH"
				StrProduct="64-bit Read-Only"
		end select	
	end if
	
	Set Conn=CreateObject("ADODB.Connection")
	Conn.Mode=adModeRead
	Set Recset=CreateObject("ADODB.Recordset")
	Conn.Open "Driver=SQL Server;Server=10.6.208.62;Database=DSI;uid=sa;pwd=Quest6848;"
	'wscript.echo("the Product Name is: " + StrProduct + " and the Version is: " + StrVersion)
	isSQL="Update DSI.dbo.DSI_FinishInstall_ToadforOracle set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and I_AutoUpdate = 'Ture' and I_ProductName like 'Toad% for Oracle%" + StrProduct +"'"
	
	'wscript.echo(isSQL)
	Conn.Execute "Update DSI.dbo.DSI_FinishInstall_ToadforOracle set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and I_AutoUpdate = 'Ture' and I_ProductName like 'Toad% for Oracle%" + StrProduct +"'"
	
	Conn.Close
	set Conn=Nothing
	
	if Err.Number = 0 then
		Update_DSI_FinishInstall_ToadforOracle=True
	else
		Update_DSI_FinishInstall_ToadforOracle=False
		Err.Clear
	end if

End Function

'==============================DSI_FinishInstall_BMF========================================
Function Update_DSI_FinshInstall_OptimizerforOracle(StrProduct,StrVersion)

	Dim Conn,isSQL
	on error resume next
	
	if IsEmpty(StrProduct) then
		Update_DSI_FinshInstall_OptimizerforOracle=false
		wscript.quit 100
	else
		select case StrProduct
			case "SQLOPTIMIZERFORORACLE_X64_MULTILANG"
				StrProduct="64-bit"
			case "SQLOPTIMIZERFORORACLE_X86_MULTILANG"
				StrProduct="32-bit"
			case "SQLOPTIMIZERFORORACLE_TRIAL_X86_MULTILANG"
				StrProduct="32-bit Trial"
			case "SQLOPTIMIZERFORORACLE_TRIAL_X64_MULTILANG"
				StrProduct="64-bit Trial"
		end select	
	end if
	
	Set Conn=CreateObject("ADODB.Connection")
	Conn.Mode=adModeRead
	Set Recset=CreateObject("ADODB.Recordset")
	Conn.Open "Driver=SQL Server;Server=10.6.208.62;Database=DSI;uid=sa;pwd=Quest6848;"
	'wscript.echo("the Product Name is: " + StrProduct + " and the Version is: " + StrVersion)
	isSQL="Update DSI.dbo.DSI_FinshInstall_OptimizerforOracle set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and I_AutoUpdate = 'Ture' and I_ProductName like 'Dell% SQL Optimizer for Oracle%" + StrProduct +"'"
	
	'wscript.echo(isSQL)
	Conn.Execute "Update DSI.dbo.DSI_FinshInstall_OptimizerforOracle set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and I_AutoUpdate = 'Ture' and I_ProductName like 'Dell% SQL Optimizer for Oracle%" + StrProduct +"'"
	
	Conn.Close
	set Conn=Nothing
	
	if Err.Number = 0 then
		Update_DSI_FinshInstall_OptimizerforOracle=True
	else
		Update_DSI_FinshInstall_OptimizerforOracle=False
		Err.Clear
	end if

End Function

'==============================DSI_DSI_FinishInstall_BMF========================================
Function Update_DSI_FinishInstall_BMF(StrProduct,StrVersion)

	Dim Conn,isSQL
	on error resume next
	
	if IsEmpty(StrProduct) then
		Update_DSI_FinishInstall_BMF=false
		wscript.quit 100
	else
		select case StrProduct
			case "BENCHMARKFACTORY_X64_EN"
				StrProduct="64-bit"
			case "BENCHMARKFACTORY_X86_EN"
				StrProduct="32-bit"
			case "BENCHMARKFACTORY_TRIAL_X86_EN"
				StrProduct="32-bit Trial"
			case "BENCHMARKFACTORY_TRIAL_X64_EN"
				StrProduct="64-bit Trial"
		end select	
	end if
	
	Set Conn=CreateObject("ADODB.Connection")
	Conn.Mode=adModeRead
	Set Recset=CreateObject("ADODB.Recordset")
	Conn.Open "Driver=SQL Server;Server=10.6.208.62;Database=DSI;uid=sa;pwd=Quest6848;"
	'wscript.echo("the Product Name is: " + StrProduct + " and the Version is: " + StrVersion)
	isSQL="Update DSI.dbo.DSI_FinishInstall_BMF set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and I_AutoUpdate = 'Ture' and I_ProductName like 'Benchmark Factory%" + StrProduct +"'"
	
	'wscript.echo(isSQL)
	Conn.Execute "Update DSI.dbo.DSI_FinishInstall_BMF set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and I_AutoUpdate = 'Ture' and I_ProductName like 'Benchmark Factory%" + StrProduct +"'"
	
	Conn.Close
	set Conn=Nothing
	
	if Err.Number = 0 then
		Update_DSI_FinishInstall_BMF=True
	else
		Update_DSI_FinishInstall_BMF=False
		Err.Clear
	end if

End Function

'==============================DSI_FinishInstall_SpotlightonOracle========================================
Function Update_DSI_FinishInstall_SpotlightonOracle(StrProduct,StrVersion)

	Dim Conn,isSQL
	on error resume next
	
	if IsEmpty(StrProduct) then
		Update_DSI_FinishInstall_SpotlightonOracle=false
		wscript.quit 100
	else
		select case StrProduct
			case "SPOTLIGHTONORACLE_X64_MULTILANG"
				StrProduct="64-bit"
			case "SPOTLIGHTONORACLE_X86_MULTILANG"
				StrProduct="32-bit"
		end select	
	end if
	
	Set Conn=CreateObject("ADODB.Connection")
	Conn.Mode=adModeRead
	Set Recset=CreateObject("ADODB.Recordset")
	Conn.Open "Driver=SQL Server;Server=10.6.208.62;Database=DSI;uid=sa;pwd=Quest6848;"
	'wscript.echo("the Product Name is: " + StrProduct + " and the Version is: " + StrVersion)
	isSQL="Update DSI.dbo.DSI_FinishInstall_SpotlightonOracle set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and I_AutoUpdate = 'Ture' and I_ProductName like 'Spotlight% on Oracle%" + StrProduct +"'"
	
	'wscript.echo(isSQL)
	Conn.Execute "Update DSI.dbo.DSI_FinishInstall_SpotlightonOracle set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and I_AutoUpdate = 'Ture' and I_ProductName like 'Spotlight% on Oracle%" + StrProduct +"'"
	
	Conn.Close
	set Conn=Nothing
	
	if Err.Number = 0 then
		Update_DSI_FinishInstall_SpotlightonOracle=True
	else
		Update_DSI_FinishInstall_SpotlightonOracle=False
		Err.Clear
	end if

End Function

'==============================DSI_FinishInstall_ToadDataModeler========================================
Function Update_DSI_FinishInstall_ToadDataModeler(StrProduct,StrVersion)

	Dim Conn,isSQL
	on error resume next
	
	if IsEmpty(StrProduct) then
		Update_DSI_FinishInstall_ToadDataModeler=false
		wscript.quit 100
	else
		select case StrProduct
			case "TOADDATAMODELER_X86_EN"
				StrProduct="32-bit"
			case "TOADDATAMODELER_X64_EN"
				StrProduct="64-bit"
		end select	
	end if
	
	Set Conn=CreateObject("ADODB.Connection")
	Conn.Mode=adModeRead
	Set Recset=CreateObject("ADODB.Recordset")
	Conn.Open "Driver=SQL Server;Server=10.6.208.62;Database=DSI;uid=sa;pwd=Quest6848;"	
	
	Conn.Execute "Update DSI.dbo.DSI_FinishInstall_ToadDataModeler set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TURE' and I_ProductName like 'Toad% Data Modeler'"
	
	Conn.Close
	set Conn=Nothing
	
	if Err.Number = 0 then
		Update_DSI_FinishInstall_ToadDataModeler=True
	else
		Update_DSI_FinishInstall_ToadDataModeler=False
		Err.Clear
	end if

End Function

'==============================DSI_FinishInstall_QuestCodeTester========================================
Function Update_DSI_FinishInstall_QuestCodeTester(StrProduct,StrVersion)

	Dim Conn,isSQL
	on error resume next
	
	if IsEmpty(StrProduct) then
		Update_DSI_FinishInstall_QuestCodeTester=false
		wscript.quit 100
	else
		select case StrProduct
			case "CODETESTERORACLE_X86_EN"
				StrProduct="32-bit"
			case "CODETESTERORACLE_X64_EN"
				StrProduct="64-bit"
		end select	
	end if
	
	Set Conn=CreateObject("ADODB.Connection")
	Conn.Mode=adModeRead
	Set Recset=CreateObject("ADODB.Recordset")
	Conn.Open "Driver=SQL Server;Server=10.6.208.62;Database=DSI;uid=sa;pwd=Quest6848;"	
	
	Conn.Execute "Update DSI.dbo.DSI_FinishInstall_QuestCodeTester set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TURE' and UPPER(I_ProductName) like 'DELL% CODE TESTER FOR ORACLE'"
	
	Conn.Close
	set Conn=Nothing
	
	if Err.Number = 0 then
		Update_DSI_FinishInstall_QuestCodeTester=True
	else
		Update_DSI_FinishInstall_QuestCodeTester=False
		Err.Clear
	end if

End Function

'==============================DSI_FinishInstall_BackupReportForOracle========================================
Function Update_DSI_FinishInstall_BackupReportForOracle(StrProduct,StrVersion)

	Dim Conn,isSQL
	on error resume next
	
	if IsEmpty(StrProduct) then
		Update_DSI_FinishInstall_BackupReportForOracle=false
		wscript.quit 100
	else
		select case StrProduct
			case "BACKUPREPORTER_X86_EN"
				StrProduct="32-bit"
			case "BACKUPREPORTER_X64_EN"
				StrProduct="64-bit"
		end select	
	end if
	
	Set Conn=CreateObject("ADODB.Connection")
	Conn.Mode=adModeRead
	Set Recset=CreateObject("ADODB.Recordset")
	Conn.Open "Driver=SQL Server;Server=10.6.208.62;Database=DSI;uid=sa;pwd=Quest6848;"	
	
	Conn.Execute "Update DSI.dbo.DSI_FinishInstall_BackupReportForOracle set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TURE' and UPPER(I_ProductName) like 'DELL% BACKUP REPORTER FOR ORACLE'"
	
	Conn.Close
	set Conn=Nothing
	
	if Err.Number = 0 then
		Update_DSI_FinishInstall_BackupReportForOracle=True
	else
		Update_DSI_FinishInstall_BackupReportForOracle=False
		Err.Clear
	end if

End Function


'==============================DSI_FinishInstall_BackupReportForOracle========================================
Function Update_DSI_FinishInstall_ToadforMySQLFreeware(StrProduct,StrVersion)

	Dim Conn,isSQL
	on error resume next
	
	if IsEmpty(StrProduct) then
		Update_DSI_FinishInstall_ToadforMySQLFreeware=false
		wscript.quit 100
	else
		select case StrProduct
			case "TOADFORMYSQL_FREEWARE_X86_EN"
				StrProduct="32-bit"
			case "TOADFORMYSQL_FREEWARE_X64_EN"
				StrProduct="64-bit"
		end select	
	end if
	
	Set Conn=CreateObject("ADODB.Connection")
	Conn.Mode=adModeRead
	Set Recset=CreateObject("ADODB.Recordset")
	Conn.Open "Driver=SQL Server;Server=10.6.208.62;Database=DSI;uid=sa;pwd=Quest6848;"	
	
	Conn.Execute "Update DSI.dbo.DSI_FinishInstall_ToadforMySQLFreeware set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TURE' and UPPER(I_ProductName) like 'TOAD% FOR MYSQL'"
	
	Conn.Close
	set Conn=Nothing
	
	if Err.Number = 0 then
		Update_DSI_FinishInstall_ToadforMySQLFreeware=True
	else
		Update_DSI_FinishInstall_ToadforMySQLFreeware=False
		Err.Clear
	end if

End Function


'================================================================================

Sub UpdateTestData()
	on error resume next

	Dim XMLDoc
	Dim ErrorMsg
	Dim Position
	Dim ProjectFile
	Dim productName,productversion,StrProduct
	Dim ParentGroup,groupowner,childgroup

	If WScript.Arguments.Count=2 then
		ProjectFile=Trim(WScript.Arguments(0))
		StrProduct=Trim(WScript.Arguments(1))
	Else
		wscript.Quit 400
	End If
	

	set XMLDoc=CreateObject("MSXML2.DOMDOCUMENT")

	XMLDoc.async=False

	XMLDoc.ValidateonParse=True
	'Open project file
	Set FSO=CreateObject("Scripting.FileSystemObject")
	if not FSO.FileExists(ProjectFile) then
		set FSO=Nothing
		wscript.quit 404
	end if
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
	Set ProductNode=XMLDOC.selectNodes("//Include")
	
	'wscript.echo("there are total: " + cstr(RootNode.childnodes.length) + " Products in this build")
	
	set ProductNode=Productnode.item(0)
	set regEx= New RegExp
	
	For i = 0 to RootNode.childnodes.length - 1
		
			NodeName=productnode.childnodes.item(i).text
			NodeName=Split(NodeName,"=")
			
			regEx.Pattern = "\d+(\.\d+)+"
			regEx.Global=True
			set Matches =regEx.Execute(NodeName(1))
			
			for each match in Matches
				ProductVersion=match.value
			next
			ProductName=Mid(NodeName(0),1,Len(NodeName(0)) - 15)
		
			if ProductName <> "" and ProductVersion <> "" then
				if InStr(ProductName,"TOADFORORACLE") = 1 then
					if Update_DSI_FinishInstall_ToadforOracle(ProductName,ProductVersion) then
						'wscript.echo("Update DSI_FinishInstall_ToadforOracle table successful!")
					end if
				end if
				if InStr(ProductName,"SQLOPTIMIZERFORORACLE") = 1 then
					if Update_DSI_FinshInstall_OptimizerforOracle(ProductName,ProductVersion) then
						'wscript.echo("Update Update_DSI_FinshInstall_OptimizerforOracle table successful!")
					end if
				end if
				if InStr(ProductName,"BENCHMARKFACTORY") = 1 then
					if Update_DSI_FinishInstall_BMF(ProductName,ProductVersion) then
						'wscript.echo("Update Update_DSI_FinishInstall_BMF table successful!")
					end if
				end if
				if InStr(ProductName,"SPOTLIGHTONORACLE") = 1 then
					if Update_DSI_FinishInstall_SpotlightonOracle(ProductName,ProductVersion) then
						'wscript.echo("Update Update_DSI_FinishInstall_SpotlightonOracle table successful!")
					end if
				end if
				if InStr(ProductName,"TOADDATAMODELER") = 1 then
					if Update_DSI_FinishInstall_ToadDataModeler(ProductName,ProductVersion) then
						'wscript.echo("Update Update_DSI_FinishInstall_ToadDataModeler table successful!")
					end if
				end if
				if InStr(UCase(ProductName),UCase("CODETESTERORACLE")) = 1 then
					if Update_DSI_FinishInstall_QuestCodeTester(ProductName,ProductVersion) then
						'wscript.echo("Update Update_DSI_FinishInstall_QuestCodeTester table successful!")
					end if
				end if
				if InStr(UCase(ProductName),UCase("BACKUPREPORTER")) = 1 then
					if Update_DSI_FinishInstall_BackupReportForOracle(ProductName,ProductVersion) then
						'wscript.echo("Update Update_DSI_FinishInstall_BackupReportForOracle table successful!")
					end if
				end if
				if InStr(UCase(ProductName),UCase("TOADFORMYSQL")) = 1 then
					if Update_DSI_FinishInstall_ToadforMySQLFreeware(ProductName,ProductVersion) then
						'wscript.echo("Update Update_DSI_FinishInstall_ToadforMySQLFreeware table successful!")
					end if
				end if
			end if
	Next
	
		
End Sub

Call UpdateTestData()

