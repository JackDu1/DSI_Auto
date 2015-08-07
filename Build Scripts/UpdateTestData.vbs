Dim Conn

Class UpdateOracleSuite

	'==============================DSI_FinishInstall_ToadforOracle========================================
	Function Update_DSI_FinishInstall_ToadforOracle(StrProduct,StrVersion)

		on error resume next
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_ToadforOracle=false
			wscript.quit 100
		else
			select case UCase(StrProduct)
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
				case else
					StrProduct=""
			end select	
		end if

		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_ToadforOracle set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TURE' and I_ProductName like 'Toad% for Oracle%" + StrProduct +"'"

		if Err.Number = 0 then
			Update_DSI_FinishInstall_ToadforOracle=True
		else
			Update_DSI_FinishInstall_ToadforOracle=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinshInstall_OptimizerforOracle========================================
	Function Update_DSI_FinshInstall_OptimizerforOracle(StrProduct,StrVersion)

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
		
		Conn.Execute "Update DSI.dbo.DSI_FinshInstall_OptimizerforOracle set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and I_AutoUpdate = 'Ture' and I_ProductName like 'Dell% SQL Optimizer for Oracle%" + StrProduct +"'"
		
		if Err.Number = 0 then
			Update_DSI_FinshInstall_OptimizerforOracle=True
		else
			Update_DSI_FinshInstall_OptimizerforOracle=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_BMF========================================
	Function Update_DSI_FinishInstall_BMF(StrProduct,StrVersion)

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

		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_BMF set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and I_AutoUpdate = 'Ture' and I_ProductName like 'Benchmark Factory%" + StrProduct +"'"

		if Err.Number = 0 then
			Update_DSI_FinishInstall_BMF=True
		else
			Update_DSI_FinishInstall_BMF=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_SpotlightonOracle========================================
	Function Update_DSI_FinishInstall_SpotlightonOracle(StrProduct,StrVersion)

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
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_SpotlightonOracle set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and I_AutoUpdate = 'Ture' and I_ProductName like 'Spotlight% on Oracle%" + StrProduct +"'"

		if Err.Number = 0 then
			Update_DSI_FinishInstall_SpotlightonOracle=True
		else
			Update_DSI_FinishInstall_SpotlightonOracle=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_ToadDataModeler========================================
	Function Update_DSI_FinishInstall_ToadDataModeler(StrProduct,StrVersion)

		on error resume next
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_ToadDataModeler=false
			wscript.quit 100
		end if
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_ToadDataModeler set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TURE' and I_ProductName like 'Toad% Data Modeler'"
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_ToadDataModeler=True
		else
			Update_DSI_FinishInstall_ToadDataModeler=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_QuestCodeTester========================================
	Function Update_DSI_FinishInstall_QuestCodeTester(StrProduct,StrVersion)

		on error resume next
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_QuestCodeTester=false
			wscript.quit 100
		end if
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_QuestCodeTester set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TURE' and UPPER(I_ProductName) like 'DELL% CODE TESTER FOR ORACLE'"
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_QuestCodeTester=True
		else
			Update_DSI_FinishInstall_QuestCodeTester=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_BackupReportForOracle========================================
	Function Update_DSI_FinishInstall_BackupReportForOracle(StrProduct,StrVersion)

		on error resume next
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_BackupReportForOracle=false
			wscript.quit 100	
		end if
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_BackupReportForOracle set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TURE' and UPPER(I_ProductName) like 'DELL% BACKUP REPORTER FOR ORACLE'"
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_BackupReportForOracle=True
		else
			Update_DSI_FinishInstall_BackupReportForOracle=False
			Err.Clear
		end if

	End Function


	'==============================DSI_FinishInstall_BackupReportForOracle========================================
	Function Update_DSI_FinishInstall_ToadforMySQLFreeware(StrProduct,StrVersion)

		on error resume next
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_ToadforMySQLFreeware=false
			wscript.quit 100	
		end if
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_ToadforMySQLFreeware set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TURE' and UPPER(I_ProductName) like 'TOAD% FOR MYSQL'"
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_ToadforMySQLFreeware=True
		else
			Update_DSI_FinishInstall_ToadforMySQLFreeware=False
			Err.Clear
		end if

	End Function

End Class

Class UpdateSAPSuite

	'==============================DSI_FinishInstall_ToadforSybase========================================
	Function Update_DSI_FinishInstall_ToadforSybase(StrProduct,StrVersion)

		on error resume next
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_ToadforSybase=false
			wscript.quit 100	
		end if

		Conn.Execute "Update DSI.dbo.DSI_SAP_ToadforSybase set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TURE' and I_ProductName like 'Toad% for SAP Solutions%'"

		if Err.Number = 0 then
			Update_DSI_FinishInstall_ToadforSybase=True
		else
			Update_DSI_FinishInstall_ToadforSybase=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_QuestSQLOptimizerforSybase========================================
	Function Update_DSI_FinishInstall_QuestSQLOptimizerforSybase(StrProduct,StrVersion)

		on error resume next
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinshInstall_OptimizerforOracle=false
			wscript.quit 100
		end if
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_QuestSQLOptimizerforSybase set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TURE' and UPPER(I_ProductName) like 'DELL% SQL OPTIMIZER FOR SAP% ASE'"
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_QuestSQLOptimizerforSybase=True
		else
			Update_DSI_FinishInstall_QuestSQLOptimizerforSybase=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_BMF========================================
	Function Update_DSI_FinishInstall_BMF(StrProduct,StrVersion)

		on error resume next
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_BMF=false
			wscript.quit 100	
		else
			select case StrProduct
				case "BENCHMARKFACTORY_X64_EN"
					StrProduct="for Databases"
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="for Databases"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="for Databases Trial"
				case "BENCHMARKFACTORY_TRIAL_X64_EN"
					StrProduct="for Databases Trial"
			end select
		end if

		Conn.Execute "Update DSI.dbo.DSI_SAP_BMF set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TURE' and I_ProductName like 'Benchmark Factory%" + StrProduct + "'"

		if Err.Number = 0 then
			Update_DSI_FinishInstall_BMF=True
		else
			Update_DSI_FinishInstall_BMF=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_SpotlightonSybase========================================
	Function Update_DSI_FinishInstall_SpotlightonSybase(StrProduct,StrVersion)

		on error resume next
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_SpotlightonSybase=false
			wscript.quit 100	
		end if
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_SpotlightonSybase set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TURE' and I_ProductName like 'Spotlight% on SAP% ASE'"

		if Err.Number = 0 then
			Update_DSI_FinishInstall_SpotlightonSybase=True
		else
			Update_DSI_FinishInstall_SpotlightonSybase=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_ToadDataModeler========================================
	Function Update_DSI_FinishInstall_ToadDataModeler(StrProduct,StrVersion)

		on error resume next
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_ToadDataModeler=false
			wscript.quit 100
		end if
		
		Conn.Execute "Update DSI.dbo.DSI_SAP_ToadDataModeler set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TURE' and I_ProductName like 'Toad% Data Modeler'"
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_ToadDataModeler=True
		else
			Update_DSI_FinishInstall_ToadDataModeler=False
			Err.Clear
		end if

	End Function

End Class

Class UpdateDB2Suite

	'==============================DSI_FinishInstall_ToadforIBMDB2LUW========================================
	Function Update_DSI_FinishInstall_ToadforIBMDB2LUW(StrProduct,StrVersion)

		'on error resume next
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_ToadforIBMDB2LUW=false
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "TOADFORDB2_X86_EN"
					StrProduct=""
				case "TOADFORDB2_TRIAL_X86_EN"
					StrProduct="Trial"
			end select	
		end if

		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_ToadforIBMDB2LUW set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TURE' and I_ProductName like 'Toad% for IBM% DB2%" + StrProduct +"'"

		if Err.Number = 0 then
			Update_DSI_FinishInstall_ToadforIBMDB2LUW=True
		else
			Update_DSI_FinishInstall_ToadforIBMDB2LUW=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_QuestSQLOptimizerforIBMDB2========================================
	Function Update_DSI_FinishInstall_QuestSQLOptimizerforIBMDB2(StrProduct,StrVersion)

		on error resume next
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_QuestSQLOptimizerforIBMDB2=false
			wscript.quit 100	
		end if
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_QuestSQLOptimizerforIBMDB2 set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 2 and I_AutoUpdate = 'Ture' and I_ProductName like 'Dell% SQL Optimizer for IBM% DB2% LUW'"
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_QuestSQLOptimizerforIBMDB2=True
		else
			Update_DSI_FinishInstall_QuestSQLOptimizerforIBMDB2=False
			Err.Clear
		end if

	End Function
	
	'==============================DSI_FinishInstall_QuestSQLOptimizerForDB2zOS========================================
	Function Update_DSI_FinishInstall_QuestSQLOptimizerForDB2zOS(StrProduct,StrVersion)

		on error resume next
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_QuestSQLOptimizerForDB2zOS=false
			wscript.quit 100	
		end if
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_QuestSQLOptimizerForDB2zOS set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 2 and I_AutoUpdate = 'Ture' and I_ProductName like 'Dell% SQL Optimizer for IBM% DB2% z/OS%'"
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_QuestSQLOptimizerForDB2zOS=True
		else
			Update_DSI_FinishInstall_QuestSQLOptimizerForDB2zOS=False
			Err.Clear
		end if

	End Function

	'==============================DSI_DSI_FinishInstall_BMF========================================
	Function Update_DSI_FinishInstall_BMF(StrProduct,StrVersion)

		on error resume next
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_BMF=false
			wscript.quit 100
		else
			select case StrProduct
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="for Databases"
				case "BENCHMARKFACTORY_X64_EN"
					StrProduct="for Databases"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="for Databases Trial"
				case "BENCHMARKFACTORY_TRIAL_X64_EN"
					StrProduct="for Databases Trial"
			end select	
		end if

		Conn.Execute "Update DSI.dbo.DSI_DB2_BMF set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TURE' and I_ProductName like 'Benchmark Factory%" + StrProduct +"'"

		if Err.Number = 0 then
			Update_DSI_FinishInstall_BMF=True
		else
			Update_DSI_FinishInstall_BMF=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_SpotlightonIBMDB2========================================
	Function Update_DSI_FinishInstall_SpotlightonIBMDB2(StrProduct,StrVersion)

		on error resume next
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_SpotlightonIBMDB2=false
			wscript.quit 100	
		end if
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_SpotlightonIBMDB2 set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TURE' and I_ProductName like 'Spotlight% on IBM% DB2% LUW'"

		if Err.Number = 0 then
			Update_DSI_FinishInstall_SpotlightonIBMDB2=True
		else
			Update_DSI_FinishInstall_SpotlightonIBMDB2=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_ToadDataModeler========================================
	Function Update_DSI_FinishInstall_ToadDataModeler(StrProduct,StrVersion)

		on error resume next
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_ToadDataModeler=false
			wscript.quit 100	
		end if
		
		Conn.Execute "Update DSI.dbo.DSI_DB2_ToadDataModeler set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TURE' and I_ProductName like 'Toad% Data Modeler'"
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_ToadDataModeler=True
		else
			Update_DSI_FinishInstall_ToadDataModeler=False
			Err.Clear
		end if

	End Function

End Class

Class UpdateSQLServerSuite

	'==============================DSI_DSI_FinishInstall_BMF========================================
	Function Update_DSI_FinishInstall_BMF(StrProduct,StrVersion)

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

		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_BMF set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 4 and I_AutoUpdate = 'Ture' and I_ProductName like 'Benchmark Factory%" + StrProduct +"'"

		if Err.Number = 0 then
			Update_DSI_FinishInstall_BMF=True
		else
			Update_DSI_FinishInstall_BMF=False
			Err.Clear
		end if

	End Function


	'==============================DSI_FinishInstall_ToadDataModeler========================================
	Function Update_DSI_FinishInstall_ToadDataModeler(StrProduct,StrVersion)

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
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_ToadDataModeler set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 4 and UPPER(I_AutoUpdate) = 'TURE' and I_ProductName like 'Toad% Data Modeler'"
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_ToadDataModeler=True
		else
			Update_DSI_FinishInstall_ToadDataModeler=False
			Err.Clear
		end if

	End Function

End Class

'================================================================================

Sub UpdateTestData()
	on error resume next

	Dim XMLDoc
	Dim ErrorMsg
	Dim NewOracleSuite,NewSAPSuite,NewDB2Suite,NewSQLServerSuite
	Dim ProjectFile
	Dim productName,productversion,StrProduct,PreProduct
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
	
	Set Conn=CreateObject("ADODB.Connection")
	Conn.Mode=adModeRead
	Conn.Open "Driver=SQL Server;Server=10.6.208.62;Database=DSI;uid=sa;pwd=Quest6848;"	
	
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
		'ProductName=Split(NodeName(0),"PACKAGE")
		
		if ProductName <> "" and ProductVersion <> "" then
			PreProduct=Split(ProductName,"_")
			Select Case UCase(StrProduct)
				case UCase("ORACLE")
					Set NewOracleSuite=New UpdateOracleSuite
					Select Case Trim(UCase(PreProduct(0)))
						case "TOADFORORACLE"
							if NewOracleSuite.Update_DSI_FinishInstall_ToadforOracle(ProductName,ProductVersion) then
								wscript.echo("Update DSI_FinishInstall_ToadforOracle table successful!")
							end if
						case "SQLOPTIMIZERFORORACLE"
							if NewOracleSuite.Update_DSI_FinshInstall_OptimizerforOracle(ProductName,ProductVersion) then
								'wscript.echo("Update Update_DSI_FinshInstall_OptimizerforOracle table successful!")
							end if
						case "BENCHMARKFACTORY"
							if NewOracleSuite.Update_DSI_FinishInstall_BMF(ProductName,ProductVersion) then
								'wscript.echo("Update Update_DSI_FinishInstall_BMF table successful!")
							end if
						case "SPOTLIGHTONORACLE"
							if NewOracleSuite.Update_DSI_FinishInstall_SpotlightonOracle(ProductName,ProductVersion) then
								'wscript.echo("Update Update_DSI_FinishInstall_SpotlightonOracle table successful!")
							end if
						case "TOADDATAMODELER"
							if NewOracleSuite.Update_DSI_FinishInstall_ToadDataModeler(ProductName,ProductVersion) then
								'wscript.echo("Update Update_DSI_FinishInstall_ToadDataModeler table successful!")
							end if
						case "CODETESTERORACLE"
							if NewOracleSuite.Update_DSI_FinishInstall_QuestCodeTester(ProductName,ProductVersion) then
								'wscript.echo("Update Update_DSI_FinishInstall_QuestCodeTester table successful!")
							end if
						case "BACKUPREPORTER"
							if NewOracleSuite.Update_DSI_FinishInstall_BackupReportForOracle(ProductName,ProductVersion) then
								'wscript.echo("Update Update_DSI_FinishInstall_BackupReportForOracle table successful!")
							end if
						case "TOADFORMYSQL"
							if NewOracleSuite.Update_DSI_FinishInstall_ToadforMySQLFreeware(ProductName,ProductVersion) then
								'wscript.echo("Update Update_DSI_FinishInstall_ToadforMySQLFreeware table successful!")
							end if
					End Select
				case UCase("DB2")
					Set NewDB2Suite=New UpdateDB2Suite
					Select Case Trim(UCase(PreProduct(0)))
						case "TOADFORDB2"
							if NewDB2Suite.Update_DSI_FinishInstall_ToadforIBMDB2LUW(ProductName,ProductVersion) then
								wscript.echo("Update DSI_DSI_FinishInstall_ToadforIBMDB2LUW table successful!")
							end if
						case "SQLOPTIMIZERFORDB2LUW"
							if NewDB2Suite.Update_DSI_FinishInstall_QuestSQLOptimizerforIBMDB2(ProductName,ProductVersion) then
								'wscript.echo("Update Update_DSI_FinishInstall_QuestSQLOptimizerforIBMDB2 table successful!")
							end if
						case "SQLOPTIMIZERFORDB2ZOS"
							if NewDB2Suite.Update_DSI_FinishInstall_QuestSQLOptimizerForDB2zOS(ProductName,ProductVersion) then
								'wscript.echo("Update Update_DSI_FinishInstall_QuestSQLOptimizerForDB2zOS table successful!")
							end if
						case "BENCHMARKFACTORY"
							if NewDB2Suite.Update_DSI_FinishInstall_BMF(ProductName,ProductVersion) then
								'wscript.echo("Update Update_DSI_FinishInstall_BMF table successful!")
							end if
						case "SPOTLIGHTONDB2"
							if NewDB2Suite.Update_DSI_FinishInstall_SpotlightonIBMDB2(ProductName,ProductVersion) then
								'wscript.echo("Update Update_DSI_FinishInstall_SpotlightonIBMDB2 table successful!")
							end if
						case "TOADDATAMODELER"
							if NewDB2Suite.Update_DSI_FinishInstall_ToadDataModeler(ProductName,ProductVersion) then
								'wscript.echo("Update Update_DSI_FinishInstall_ToadDataModeler table successful!")
							end if
					End Select
				case UCase("SAP")
					Set NewSAPSuite=New UpdateSAPSuite
					Select Case Trim(UCase(PreProduct(0)))
						case "TOADFORSAP"
							if NewSAPSuite.Update_DSI_FinishInstall_ToadforSybase(ProductName,ProductVersion) then
								wscript.echo("Update DSI_FinishInstall_ToadforSybase table successful!")
							end if
						case "SQLOPTIMIZERFORSAP"
							if NewSAPSuite.Update_DSI_FinishInstall_QuestSQLOptimizerforSybase(ProductName,ProductVersion) then
								'wscript.echo("Update DSI_FinishInstall_QuestSQLOptimizerforSybase table successful!")
							end if
						case "BENCHMARKFACTORY"
							if NewSAPSuite.Update_DSI_FinishInstall_BMF(ProductName,ProductVersion) then
								'wscript.echo("Update DSI_FinishInstall_BMF table successful!")
							end if
						case "SPOTLIGHTONSAP"
							if NewSAPSuite.Update_DSI_FinishInstall_SpotlightonSybase(ProductName,ProductVersion) then
								'wscript.echo("Update DSI_FinishInstall_SpotlightonSybase table successful!")
							end if
						case "TOADDATAMODELER"
							if NewSAPSuite.Update_DSI_FinishInstall_ToadDataModeler(ProductName,ProductVersion) then
								'wscript.echo("Update DSI_FinishInstall_ToadDataModeler table successful!")
							end if
					End Select
				case UCase("SQLSERVER")
					Set NewSQLServerSuite=New UpdateSQLSERVERSuite
					'not implemented
			end Select
		end if
	Next
	
	Conn.Close
	set Conn=Nothing
	
End Sub

Call UpdateTestData()