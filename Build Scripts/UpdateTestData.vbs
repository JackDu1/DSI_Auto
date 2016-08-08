'Option Explicit

Dim Conn

Class UpdateOracleSuite

	'==============================DSI_FinishInstall_ToadforOracle========================================
	Sub Update_DSI_FinishInstall_ToadforOracle(ByVal StrProduct,ByVal StrVersion)
		
		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
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
			end select	
		end if
		'Update I_Version Column Record
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_ToadforOracle set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Toad% for Oracle%" + StrProduct +"'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_FinishInstall_ToadforOracle where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Toad% for Oracle%" + StrProduct +"'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_ToadforOracle set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Toad% for Oracle%" + StrProduct +"'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub

	'==============================DSI_FinshInstall_OptimizerforOracle========================================
	Sub Update_DSI_FinshInstall_OptimizerforOracle(ByVal StrProduct,ByVal StrVersion)
		
		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
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
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_FinshInstall_OptimizerforOracle set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Dell% SQL Optimizer for Oracle%" + StrProduct +"'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_FinshInstall_OptimizerforOracle where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Dell% SQL Optimizer for Oracle%" + StrProduct +"'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_FinshInstall_OptimizerforOracle set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Dell% SQL Optimizer for Oracle%" + StrProduct +"'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub

	'==============================DSI_FinishInstall_BMF========================================
	Sub Update_DSI_FinishInstall_BMF(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer,StrColDisplay
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case StrProduct
				case "BENCHMARKFACTORY_X64_EN"
					StrProduct="BENCHMARK FACTORY% 64-BIT"
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="BENCHMARK FACTORY% 32-BIT"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="BENCHMARK FACTORY% 32-BIT TRIAL"
				case "BENCHMARKFACTORY_TRIAL_X64_EN"
					StrProduct="BENCHMARK FACTORY% 64-BIT TRIAL"
			end select	
		end if
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_BMF set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_FinishInstall_BMF where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(2)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_BMF set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		
		'Update I_DisplayVersion Column Record
		Query		= 	"Select I_DisplayVersion from DSI.dbo.DSI_FinishInstall_BMF where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColDisplay	=	Rec.Fields("I_DisplayVersion").Value
			Rec.MoveNext
		Wend
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(3)
		if InStr(StrColDisplay,"32-bit") >= 3 then
			StrColDisplay	=	StrMainVer(0) + "." + StrMainVer(1) + " (32-bit)" + "." + StrMainVer(3)
		elseif InStr(StrColDisplay,"64-bit") >= 3  then
			StrColDisplay	=	StrMainVer(0) + "." + StrMainVer(1) + " (64-bit)" + "." + StrMainVer(3)
		else
			StrColDisplay 	= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(3)
		end if
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_BMF set  I_DisplayVersion =" + "'" + StrColDisplay + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub

	'==============================DSI_FinishInstall_SpotlightonOracle========================================
	Sub Update_DSI_FinishInstall_SpotlightonOracle(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case StrProduct
				case "SPOTLIGHTONORACLE_X64_MULTILANG"
					StrProduct="64-bit"
				case "SPOTLIGHTONORACLE_X86_MULTILANG"
					StrProduct="32-bit"
			end select	
		end if
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_SpotlightonOracle set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Spotlight% on Oracle%" + StrProduct +"'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_FinishInstall_SpotlightonOracle where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Spotlight% on Oracle%" + StrProduct +"'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_SpotlightonOracle set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Spotlight% on Oracle%" + StrProduct +"'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub

	'==============================DSI_FinishInstall_ToadDataModeler========================================
	Sub Update_DSI_FinishInstall_ToadDataModeler(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100	
		else
			select case UCase(StrProduct)
				case "TOADDATAMODELER_X86_EN"
					StrProduct="TOAD% DATA MODELER 32-BIT"
                                case "TOADDATAMODELER_X64_EN"
					StrProduct="TOAD% DATA MODELER 64-BIT"
			end select
		end if
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_ToadDataModeler set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_FinishInstall_ToadDataModeler where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_ToadDataModeler set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub

	'==============================DSI_FinishInstall_QuestCodeTester========================================
	Sub Update_DSI_FinishInstall_QuestCodeTester(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		end if
		
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_QuestCodeTester set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like 'DELL_ CODE TESTER FOR ORACLE'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_FinishInstall_QuestCodeTester where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'DELL_ CODE TESTER FOR ORACLE'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_QuestCodeTester set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'DELL_ CODE TESTER FOR ORACLE'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub

	'==============================DSI_FinishInstall_BackupReportForOracle========================================
	Sub Update_DSI_FinishInstall_BackupReportForOracle(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100	
		end if
		
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_BackupReportForOracle set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like 'DELL_ BACKUP REPORTER FOR ORACLE'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_FinishInstall_BackupReportForOracle where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'DELL_ BACKUP REPORTER FOR ORACLE'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_BackupReportForOracle set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'DELL_ BACKUP REPORTER FOR ORACLE'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub


	'==============================DSI_FinishInstall_BackupReportForOracle========================================
	Sub Update_DSI_FinishInstall_ToadforMySQLFreeware(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100	
		end if
		
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_ToadforMySQLFreeware set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like 'TOAD_ FOR MYSQL'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_FinishInstall_ToadforMySQLFreeware where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'TOAD_ FOR MYSQL'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_ToadforMySQLFreeware set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'TOAD_ FOR MYSQL'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
	
	'==============================DSI_ProductSelectionPage_VerifyProductDetail========================================
	Sub Update_DSI_ProductSelectionPage_VerifyProductDetail(ByVal StrProduct,ByVal StrVersion)
		
		on error resume next
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "TOADFORORACLE_X64_EN"
					StrProduct="TOAD%FOR ORACLE 64-BIT"
				case "TOADFORORACLE_X64_ZH"
					StrProduct="TOAD%FOR ORACLE 64-BIT"
				case "TOADFORORACLE_X86_EN"
					StrProduct="TOAD%FOR ORACLE 32-BIT"
				case "TOADFORORACLE_X86_ZH"
					StrProduct="TOAD%FOR ORACLE 32-BIT"
				case "TOADFORORACLE_TRIAL_X86_EN"
					StrProduct="TOAD%FOR ORACLE 32-BIT TRIAL"
				case "TOADFORORACLE_TRIAL_X86_ZH"
					StrProduct="TOAD%FOR ORACLE 32-BIT TRIAL"
				case "TOADFORORACLE_TRIAL_X64_EN"
					StrProduct="TOAD%FOR ORACLE 64-BIT TRIAL"
				case "TOADFORORACLE_TRIAL_X64_ZH"
					StrProduct="TOAD%FOR ORACLE 64-BIT TRIAL"
				case "TOADFORORACLE_READONLY_X86_EN"
					StrProduct="TOAD%FOR ORACLE 32-BIT READ-ONLY"
				case "TOADFORORACLE_READONLY_X86_ZH"
					StrProduct="TOAD%FOR ORACLE 32-BIT READ-ONLY"
				case "TOADFORORACLE_READONLY_X64_EN"
					StrProduct="TOAD%FOR ORACLE 64-BIT READ-ONLY"
				case "TOADFORORACLE_READONLY_X64_ZH"
					StrProduct="TOAD%FOR ORACLE 64-BIT READ-ONLY"
				case "TOADFORMYSQL_FREEWARE_X86_EN"
					StrProduct="TOAD% FOR MYSQL"
				case "BACKUPREPORTER_X86_EN"
					StrProduct="DELL% BACKUP REPORTER FOR ORACLE"
				case "CODETESTERORACLE_X86_EN"
					StrProduct="DELL% CODE TESTER FOR ORACLE"
				case "TOADDATAMODELER_X86_EN"
					StrProduct="TOAD% DATA MODELER 32-BIT"
                                case "TOADDATAMODELER_X64_EN"
					StrProduct="TOAD% DATA MODELER 64-BIT"
				case "SPOTLIGHTONORACLE_X64_MULTILANG"
					StrProduct="SPOTLIGHT% ON ORACLE 64-BIT"
				case "SPOTLIGHTONORACLE_X86_MULTILANG"
					StrProduct="SPOTLIGHT% ON ORACLE 32-BIT"
				case "BENCHMARKFACTORY_X64_EN"
					StrProduct="BENCHMARK FACTORY% 64-BIT"
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="BENCHMARK FACTORY% 32-BIT"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="BENCHMARK FACTORY% 32-BIT Trial"
				case "BENCHMARKFACTORY_TRIAL_X64_EN"
					StrProduct="BENCHMARK FACTORY% 64-BIT Trial"
				case "SQLOPTIMIZERFORORACLE_X64_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 64-BIT"
				case "SQLOPTIMIZERFORORACLE_X86_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 32-BIT"
				case "SQLOPTIMIZERFORORACLE_TRIAL_X86_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 32-BIT TRIAL"
				case "SQLOPTIMIZERFORORACLE_TRIAL_X64_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 64-BIT TRIAL"
				case else
					StrProduct="Null"
			end select
		end if
		
		Conn.Execute "Update DSI.dbo.DSI_ProductSelectionPage_VerifyProductDetail set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
	
	'==============================DSI_FinishInstall_VerifyRegistry========================================
	Sub Update_DSI_FinishInstall_VerifyRegistry(ByVal StrProduct,ByVal StrVersion)
		
		Dim StrColName,StrMainVer,Query,StrVer
		Dim Matches,match,RetStr
		
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "TOADFORORACLE_X64_EN"
					StrProduct="TOAD%FOR ORACLE 64-BIT"
				case "TOADFORORACLE_X64_ZH"
					StrProduct="TOAD%FOR ORACLE 64-BIT"
				case "TOADFORORACLE_X86_EN"
					StrProduct="TOAD%FOR ORACLE 32-BIT"
				case "TOADFORORACLE_X86_ZH"
					StrProduct="TOAD%FOR ORACLE 32-BIT"
				case "TOADFORORACLE_TRIAL_X86_EN"
					StrProduct="TOAD%FOR ORACLE 32-BIT TRIAL"
				case "TOADFORORACLE_TRIAL_X86_ZH"
					StrProduct="TOAD%FOR ORACLE 32-BIT TRIAL"
				case "TOADFORORACLE_TRIAL_X64_EN"
					StrProduct="TOAD%FOR ORACLE 64-BIT TRIAL"
				case "TOADFORORACLE_TRIAL_X64_ZH"
					StrProduct="TOAD%FOR ORACLE 64-BIT TRIAL"
				case "TOADFORORACLE_READONLY_X86_EN"
					StrProduct="TOAD%FOR ORACLE 32-BIT READ-ONLY"
				case "TOADFORORACLE_READONLY_X86_ZH"
					StrProduct="TOAD%FOR ORACLE 32-BIT READ-ONLY"
				case "TOADFORORACLE_READONLY_X64_EN"
					StrProduct="TOAD%FOR ORACLE 64-BIT READ-ONLY"
				case "TOADFORORACLE_READONLY_X64_ZH"
					StrProduct="TOAD%FOR ORACLE 64-BIT READ-ONLY"
				case "TOADFORMYSQL_FREEWARE_X86_EN"
					StrProduct="TOAD% FOR MYSQL"
				case "BACKUPREPORTER_X86_EN"
					StrProduct="DELL% BACKUP REPORTER FOR ORACLE"
				case "CODETESTERORACLE_X86_EN"
					StrProduct="DELL% CODE TESTER FOR ORACLE"
				case "TOADDATAMODELER_X86_EN"
					StrProduct="TOAD% DATA MODELER 32-BIT"
                                case "TOADDATAMODELER_X64_EN"
					StrProduct="TOAD% DATA MODELER 64-BIT"
				case "SPOTLIGHTONORACLE_X64_MULTILANG"
					StrProduct="SPOTLIGHT% ON ORACLE 64-BIT"
				case "SPOTLIGHTONORACLE_X86_MULTILANG"
					StrProduct="SPOTLIGHT% ON ORACLE 32-BIT"
				case "BENCHMARKFACTORY_X64_EN"
					StrProduct="BENCHMARK FACTORY% 64-BIT"
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="BENCHMARK FACTORY% 32-BIT"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="BENCHMARK FACTORY% 32-BIT Trial"
				case "BENCHMARKFACTORY_TRIAL_X64_EN"
					StrProduct="BENCHMARK FACTORY% 64-BIT Trial"
				case "SQLOPTIMIZERFORORACLE_X64_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 64-BIT"
				case "SQLOPTIMIZERFORORACLE_X86_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 32-BIT"
				case "SQLOPTIMIZERFORORACLE_TRIAL_X86_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 32-BIT TRIAL"
				case "SQLOPTIMIZERFORORACLE_TRIAL_X64_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 64-BIT TRIAL"
				case else
					StrProduct="Null"
			end select
		end if
		'Update I_ProductVersion Column
		Conn.Execute "Update DSI.dbo.DSI_Oracle_VerifyRegistry set  I_ProductVersion =" + "'" + StrVersion + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_InstallerDisplayProductName) like '" + StrProduct + "'"
		
		'Update I_ProductName Column
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_ProductName from DSI.dbo.DSI_Oracle_VerifyRegistry where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_InstallerDisplayProductName) like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColName=Rec.Fields("I_ProductName").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		if InStr(StrColName,"Dell Backup Reporter for Oracle") >=	1 then
			StrVer 	= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(2)
		elseif InStr(StrColName,"Benchmark Factory") >=	1 then
			StrVer 	= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(2)
		else
			StrVer 	= 	StrMainVer(0) + "." + StrMainVer(1)
		end if
		
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		Set Matches		=	RegEx.Execute(StrColName)
		For each match in matches
			RetStr		=	Match.Value
		Next
		if RetStr <> "" then
			StrColName 	= 	regEx.Replace(StrColName,StrVer)
			Conn.Execute "Update DSI.dbo.DSI_Oracle_VerifyRegistry set  I_ProductName =" + "'" + StrColName + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_InstallerDisplayProductName) like '" + StrProduct + "'"
		end if
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
	
	'==============================DSI_SilentInstallMsiBuild========================================
	Sub Update_SilentInstallMsiBuild(ByVal StrProduct,ByVal StrVersion)
		
		Dim StrColName,Query,StrVer
		Dim Matches,match,RetStr
		
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "TOADFORORACLE_X64_EN"
					StrProduct="TOAD%FOR ORACLE 64-BIT"
				case "TOADFORORACLE_X64_ZH"
					StrProduct="TOAD%FOR ORACLE 64-BIT"
				case "TOADFORORACLE_X86_EN"
					StrProduct="TOAD%FOR ORACLE 32-BIT"
				case "TOADFORORACLE_X86_ZH"
					StrProduct="TOAD%FOR ORACLE 32-BIT"
				case "TOADFORORACLE_TRIAL_X86_EN"
					StrProduct="TOAD%FOR ORACLE 32-BIT TRIAL"
				case "TOADFORORACLE_TRIAL_X86_ZH"
					StrProduct="TOAD%FOR ORACLE 32-BIT TRIAL"
				case "TOADFORORACLE_TRIAL_X64_EN"
					StrProduct="TOAD%FOR ORACLE 64-BIT TRIAL"
				case "TOADFORORACLE_TRIAL_X64_ZH"
					StrProduct="TOAD%FOR ORACLE 64-BIT TRIAL"
				case "TOADFORORACLE_READONLY_X86_EN"
					StrProduct="TOAD%FOR ORACLE 32-BIT READ-ONLY"
				case "TOADFORORACLE_READONLY_X86_ZH"
					StrProduct="TOAD%FOR ORACLE 32-BIT READ-ONLY"
				case "TOADFORORACLE_READONLY_X64_EN"
					StrProduct="TOAD%FOR ORACLE 64-BIT READ-ONLY"
				case "TOADFORORACLE_READONLY_X64_ZH"
					StrProduct="TOAD%FOR ORACLE 64-BIT READ-ONLY"
				case "TOADFORMYSQL_FREEWARE_X86_EN"
					StrProduct="TOAD% FOR MYSQL"
				case "BACKUPREPORTER_X86_EN"
					StrProduct="DELL% BACKUP REPORTER FOR ORACLE"
				case "CODETESTERORACLE_X86_EN"
					StrProduct="DELL% CODE TESTER FOR ORACLE"
				case "TOADDATAMODELER_X86_EN"
					StrProduct="TOAD% DATA MODELER 32-BIT"
                                case "TOADDATAMODELER_X64_EN"
					StrProduct="TOAD% DATA MODELER 64-BIT"
				case "SPOTLIGHTONORACLE_X64_MULTILANG"
					StrProduct="SPOTLIGHT% ON ORACLE 64-BIT"
				case "SPOTLIGHTONORACLE_X86_MULTILANG"
					StrProduct="SPOTLIGHT% ON ORACLE 32-BIT"
				case "BENCHMARKFACTORY_X64_EN"
					StrProduct="BENCHMARK FACTORY% 64-BIT"
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="BENCHMARK FACTORY% 32-BIT"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="BENCHMARK FACTORY% 32-BIT Trial"
				case "BENCHMARKFACTORY_TRIAL_X64_EN"
					StrProduct="BENCHMARK FACTORY% 64-BIT Trial"
				case "SQLOPTIMIZERFORORACLE_X64_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 64-BIT"
				case "SQLOPTIMIZERFORORACLE_X86_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 32-BIT"
				case "SQLOPTIMIZERFORORACLE_TRIAL_X86_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 32-BIT TRIAL"
				case "SQLOPTIMIZERFORORACLE_TRIAL_X64_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 64-BIT TRIAL"
				case else
					StrProduct="Null"
			end select
		end if
		
		'Update I_FilePath Column
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_FilePath from DSI.dbo.SilentInstallMsiBuild where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColName=Rec.Fields("I_FilePath").Value
			Rec.MoveNext
		Wend
		
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		Set Matches		=	RegEx.Execute(StrColName)
		For each match in matches
			RetStr		=	Match.Value
		Next
		if RetStr <> "" then
			StrColName 	= 	regEx.Replace(StrColName,StrVersion)
			Conn.Execute "Update DSI.dbo.SilentInstallMsiBuild set  I_FilePath =" + "'" + StrColName + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		end if
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
'==============================DSI_ValidateShortcutAndKeyFile========================================
	Sub Update_DSI_ValidateShortcutAndKeyFile(ByVal StrProduct,ByVal StrVersion)
		
		Dim StrColName,StrMainVer,Query,StrVer
		Dim Matches,match,RetStr
		
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "TOADFORORACLE_X64_EN"
					StrProduct="TOAD FOR ORACLE [1-9][0-9].[0-9]"
				case "TOADFORORACLE_X64_ZH"
					StrProduct="TOAD FOR ORACLE [1-9][0-9].[0-9]"
				case "TOADFORORACLE_X86_EN"
					StrProduct="TOAD FOR ORACLE [1-9][0-9].[0-9]"
				case "TOADFORORACLE_X86_ZH"
					StrProduct="TOAD FOR ORACLE [1-9][0-9].[0-9]"
				case "TOADFORORACLE_TRIAL_X86_EN"
					StrProduct="TOAD FOR ORACLE [1-9][0-9].[0-9] TRIAL"
				case "TOADFORORACLE_TRIAL_X86_ZH"
					StrProduct="TOAD FOR ORACLE [1-9][0-9].[0-9] TRIAL"
				case "TOADFORORACLE_TRIAL_X64_EN"
					StrProduct="TOAD FOR ORACLE [1-9][0-9].[0-9] TRIAL"
				case "TOADFORORACLE_TRIAL_X64_ZH"
					StrProduct="TOAD FOR ORACLE [1-9][0-9].[0-9] TRIAL"
				case "TOADFORORACLE_READONLY_X86_EN"
					StrProduct="TOAD FOR ORACLE [1-9][0-9].[0-9]"
				case "TOADFORORACLE_READONLY_X86_ZH"
					StrProduct="TOAD FOR ORACLE [1-9][0-9].[0-9]"
				case "TOADFORORACLE_READONLY_X64_EN"
					StrProduct="TOAD FOR ORACLE [1-9][0-9].[0-9]"
				case "TOADFORORACLE_READONLY_X64_ZH"
					StrProduct="TOAD FOR ORACLE [1-9][0-9].[0-9]"
				case "TOADFORMYSQL_FREEWARE_X86_EN"
					StrProduct="TOAD FOR MYSQL%"
				case "CODETESTERORACLE_X86_EN"
					StrProduct="DELL CODE TESTER FOR ORACLE%"
				case "TOADDATAMODELER_X86_EN"
					StrProduct="TOAD DATA MODELER%"
                                case "TOADDATAMODELER_X64_EN"
					StrProduct="TOAD% DATA MODELER%"
				case "SPOTLIGHTONORACLE_X64_MULTILANG"
					StrProduct="SPOTLIGHT ON ORACLE%"
				case "SPOTLIGHTONORACLE_X86_MULTILANG"
					StrProduct="SPOTLIGHT ON ORACLE%"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="BENCHMARK FACTORY TRIAL [1-9].[0-9].[0-9]"
				case "BENCHMARKFACTORY_TRIAL_X64_EN"
					StrProduct="BENCHMARK FACTORY TRIAL [1-9].[0-9].[0-9](64-bit)"
                                case "BENCHMARKFACTORY_X64_EN"					
                                        StrProduct="BENCHMARK FACTORY [1-9].[0-9].[0-9](64-bit)"				
                                case "BENCHMARKFACTORY_X86_EN"
                                     StrProduct="BENCHMARK FACTORY [1-9].[0-9].[0-9]"
				case "SQLOPTIMIZERFORORACLE_X64_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 64-BIT"
				case else
					StrProduct="Null"
			end select
		end if

		
		'Update I_ProductName Column
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_ProductName from DSI.dbo.DSI_ValidateShortcutAndKeyFile where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColName=Rec.Fields("I_ProductName").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")

		if (InStr(StrColName,"Benchmark Factory") >=	1) or (InStr(StrColName,"Benchmark Factory Trial") >=	1) then
			StrVer 	= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(2)
		else
			StrVer 	= 	StrMainVer(0) + "." + StrMainVer(1)
		end if
		
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		Set Matches		=	RegEx.Execute(StrColName)
		For each match in matches
			RetStr		=	Match.Value
		Next
		if RetStr <> "" then
			StrColName 	= 	regEx.Replace(StrColName,StrVer)
			Conn.Execute "Update DSI.dbo.DSI_ValidateShortcutAndKeyFile set  I_ProductName =" + "'" + StrColName + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		end if
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub

End Class

Class UpdateSAPSuite

	'==============================DSI_FinishInstall_ToadforSybase========================================
	Sub Update_DSI_FinishInstall_ToadforSybase(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100	
		end if
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_SAP_ToadforSybase set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Toad_ for SAP Solutions'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_SAP_ToadforSybase where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Toad_ for SAP Solutions'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_SAP_ToadforSybase set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Toad_ for SAP Solutions'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub

	'==============================DSI_FinishInstall_QuestSQLOptimizerforSybase========================================
	Sub Update_DSI_FinishInstall_QuestSQLOptimizerforSybase(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		end if
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_QuestSQLOptimizerforSybase set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like 'DELL% SQL OPTIMIZER FOR SAP% ASE'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_FinishInstall_QuestSQLOptimizerforSybase where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'DELL% SQL OPTIMIZER FOR SAP% ASE'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_QuestSQLOptimizerforSybase set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'DELL% SQL OPTIMIZER FOR SAP% ASE'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub

	'==============================DSI_FinishInstall_BMF========================================
	Sub Update_DSI_FinishInstall_BMF(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer,StrColDisplay
		on error resume next
		Set regEx = New RegExp
		if IsEmpty(StrProduct) then
			wscript.quit 100	
		else
			select case StrProduct
				case "BENCHMARKFACTORY_X64_EN"
					StrProduct="Benchmark Factory% for Databases"
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="Benchmark Factory% for Databases"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="Benchmark Factory% for Databases Trial"
				case "BENCHMARKFACTORY_TRIAL_X64_EN"
					StrProduct="Benchmark Factory% for Databases Trial"
			end select
		end if
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_SAP_BMF set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_SAP_BMF where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(2)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_SAP_BMF set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		
		'Update I_DisplayVersion Column Record
		Query		= 	"Select I_DisplayVersion from DSI.dbo.DSI_SAP_BMF where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColDisplay	=	Rec.Fields("I_DisplayVersion").Value
			Rec.MoveNext
		Wend
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(3)
		if InStr(StrColDisplay,"32-bit") >= 3 then
			StrColDisplay	=	StrMainVer(0) + "." + StrMainVer(1) + " (32-bit)" + "." + StrMainVer(3)
		elseif InStr(StrColDisplay,"64-bit") >= 3  then
			StrColDisplay	=	StrMainVer(0) + "." + StrMainVer(1) + " (64-bit)" + "." + StrMainVer(3)
		else
			StrColDisplay 	= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(3)
		end if
		
		Conn.Execute "Update DSI.dbo.DSI_SAP_BMF set  I_DisplayVersion =" + "'" + StrColDisplay + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub

	'==============================DSI_FinishInstall_SpotlightonSybase========================================
	Sub Update_DSI_FinishInstall_SpotlightonSybase(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100	
		end if
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_SpotlightonSybase set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Spotlight% on SAP% ASE'"

		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_FinishInstall_SpotlightonSybase where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Spotlight% on SAP% ASE'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_SpotlightonSybase set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Spotlight% on SAP% ASE'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub

	'==============================DSI_FinishInstall_ToadDataModeler========================================
	Sub Update_DSI_FinishInstall_ToadDataModeler(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100	
		else
			select case UCase(StrProduct)
				case "TOADDATAMODELER_X86_EN"
					StrProduct="TOAD% DATA MODELER 32-BIT"
                                case "TOADDATAMODELER_X64_EN"
					StrProduct="TOAD% DATA MODELER 64-BIT"
			end select
		end if
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_SAP_ToadDataModeler set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_SAP_ToadDataModeler where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_SAP_ToadDataModeler set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
	
	'==============================DSI_ProductSelectionPage_VerifyProductDetails========================================
	Sub Update_DSI_ProductSelectionPage_VerifyProductDetails(ByVal StrProduct,ByVal StrVersion)
	
		
				
		on error resume next
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "TOADFORSAP_X86_EN"
					StrProduct="TOAD_ FOR SAP SOLUTIONS"
				case "TOADDATAMODELER_X86_EN"
					StrProduct="TOAD% DATA MODELER 32-BIT"
                                case "TOADDATAMODELER_X64_EN"
					StrProduct="TOAD% DATA MODELER 64-BIT"
				case "SPOTLIGHTONSAP_X86_EN"
					StrProduct="SPOTLIGHT_ ON SAP_ ASE"
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="BENCHMARK FACTORY_ FOR DATABASES"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="BENCHMARK FACTORY_ FOR DATABASES TRIAL"
				case "SQLOPTIMIZERFORSAP_X86_EN"
					StrProduct="Dell_ SQL OPTIMIZER FOR SAP_ ASE"
				case else
					StrProduct="Null"
			end select
		end if
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_SAP_VerifyProductDetails set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
	
	'==============================DSI_FinishInstall_VerifyRegistry========================================
	Sub Update_DSI_FinishInstall_VerifyRegistry(ByVal StrProduct,ByVal StrVersion)
		
		Dim StrColName,StrMainVer,Query,StrVer
		Dim Matches,match,RetStr
		
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "TOADFORSAP_X86_EN"
					StrProduct="TOAD_ FOR SAP SOLUTIONS"
				case "TOADDATAMODELER_X86_EN"
					StrProduct="TOAD% DATA MODELER 32-BIT"
                                case "TOADDATAMODELER_X64_EN"
					StrProduct="TOAD% DATA MODELER 64-BIT"
				case "SPOTLIGHTONSAP_X86_EN"
					StrProduct="SPOTLIGHT_ ON SAP_ ASE"
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="BENCHMARK FACTORY_ FOR DATABASES"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="BENCHMARK FACTORY_ FOR DATABASES TRIAL"
				case "SQLOPTIMIZERFORSAP_X86_EN"
					StrProduct="Dell_ SQL OPTIMIZER FOR SAP_ ASE"
				case else
					StrProduct="Null"
			end select
		end if
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_SAP_VerifyRegistry set  I_ProductVersion =" + "'" + StrVersion + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_InstallerDisplayProductName) like '" + StrProduct + "'"
		'Update I_ProductName Column
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_ProductName from DSI.dbo.DSI_SAP_VerifyRegistry where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_InstallerDisplayProductName) like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColName=Rec.Fields("I_ProductName").Value
			Rec.MoveNext
		Wend
		StrMainVer 	= 	Split(StrVersion,".")
		if InStr(StrColName,"Benchmark Factory") >=	1 then
			StrVer 	= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(2)
		else
			StrVer 	= 	StrMainVer(0) + "." + StrMainVer(1)
		end if
		
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		Set Matches		=	RegEx.Execute(StrColName)
		For each match in matches
			RetStr		=	Match.Value
		Next
		if RetStr <> "" then
			StrColName 	= 	regEx.Replace(StrColName,StrVer)
			Conn.Execute "Update DSI.dbo.DSI_SAP_VerifyRegistry set  I_ProductName =" + "'" + StrColName + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_InstallerDisplayProductName) like '" + StrProduct + "'"
		end if
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
	
	'==============================DSI_SilentInstallMsiBuild========================================
	Sub Update_SilentInstallMsiBuild(ByVal StrProduct,ByVal StrVersion)
		
		Dim StrColName,Query,StrVer
		Dim Matches,match,RetStr
		
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "TOADFORSAP_X86_EN"
					StrProduct="TOAD_ FOR SAP SOLUTIONS"
				case "TOADDATAMODELER_X86_EN"
					StrProduct="TOAD% DATA MODELER 32-BIT"
                                case "TOADDATAMODELER_X64_EN"
					StrProduct="TOAD% DATA MODELER 64-BIT"
				case "SPOTLIGHTONSAP_X86_EN"
					StrProduct="SPOTLIGHT_ ON SAP_ ASE"
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="BENCHMARK FACTORY_ FOR DATABASES"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="BENCHMARK FACTORY_ FOR DATABASES TRIAL"
				case "SQLOPTIMIZERFORSAP_X86_EN"
					StrProduct="Dell_ SQL OPTIMIZER FOR SAP_ ASE"
				case else
					StrProduct="Null"
			end select
		end if
		
		'Update I_FilePath Column
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_FilePath from DSI.dbo.DSI_SAP_SilentInstallMsiBuild where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColName=Rec.Fields("I_FilePath").Value
			Rec.MoveNext
		Wend
		
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		Set Matches		=	RegEx.Execute(StrColName)
		For each match in matches
			RetStr		=	Match.Value
		Next
		if RetStr <> "" then
			StrColName 	= 	regEx.Replace(StrColName,StrVersion)
			Conn.Execute "Update DSI.dbo.DSI_SAP_SilentInstallMsiBuild set  I_FilePath =" + "'" + StrColName + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		end if
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
'==============================DSI_ValidateShortcutAndKeyFile========================================
	Sub Update_DSI_FinishInstall_ValidateShortcut(ByVal StrProduct,ByVal StrVersion)
		
		Dim StrColName,StrMainVer,Query,StrVer
		Dim Matches,match,RetStr
		
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "TOADFORSAP_X86_EN"
					StrProduct="TOAD_ FOR SAP SOLUTIONS%"
				case "TOADDATAMODELER_X86_EN"
					StrProduct="TOAD DATA MODELER%"
                                case "TOADDATAMODELER_X64_EN"
					StrProduct="TOAD DATA MODELER%"
				case "SPOTLIGHTONSAP_X86_EN"
					StrProduct="SPOTLIGHT%"
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="BENCHMARK FACTORY%"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="BENCHMARK FACTORY%"
				case else
					StrProduct="Null"
			end select
		end if

		
		'Update I_ProductName Column
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_ProductName from DSI.dbo.DSI_FinishInstall_ValidateShortcut where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColName=Rec.Fields("I_ProductName").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")

		if InStr(StrColName,"Benchmark Factory") >=	1 then
			StrVer 	= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(2)
		else
			StrVer 	= 	StrMainVer(0) + "." + StrMainVer(1)
		end if
		
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		Set Matches		=	RegEx.Execute(StrColName)
		For each match in matches
			RetStr		=	Match.Value
		Next
		if RetStr <> "" then
			StrColName 	= 	regEx.Replace(StrColName,StrVer)
			Conn.Execute "Update DSI.dbo.DSI_FinishInstall_ValidateShortcut set  I_ProductName =" + "'" + StrColName + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		end if
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub

End Class

Class UpdateDB2Suite

	'==============================DSI_FinishInstall_ToadforIBMDB2LUW========================================
	Sub Update_DSI_FinishInstall_ToadforIBMDB2LUW(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "TOADFORDB2_X86_EN"
					StrProduct="Toad_ for IBM_ DB2_"
				case "TOADFORDB2_TRIAL_X86_EN"
					StrProduct="Toad_ for IBM_ DB2_ Trial"
			end select	
		end if
		
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_ToadforIBMDB2LUW set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct +"'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_FinishInstall_ToadforIBMDB2LUW where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_ToadforIBMDB2LUW set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" +  StrProduct + "'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub

	'==============================DSI_FinishInstall_QuestSQLOptimizerforIBMDB2========================================
	Sub Update_DSI_FinishInstall_QuestSQLOptimizerforIBMDB2(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "SQLOPTIMIZERFORDB2LUW_X86_EN"
					StrProduct="Dell% SQL Optimizer for IBM% DB2% LUW"
				case "SQLOPTIMIZERFORDB2LUW_X64_EN"
					StrProduct="Dell% SQL Optimizer for IBM% DB2% LUW"
			end select
		end if
		
		'Update I_Version Column
		
		Conn.Execute "Update DSI.dbo.DSI_DB2_QuestSQLOptimizerforIBMDB2 set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		
		'Update I_SubFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_SubFolder from DSI.dbo.DSI_DB2_QuestSQLOptimizerforIBMDB2 where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_SubFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_DB2_QuestSQLOptimizerforIBMDB2 set  I_SubFolder =" + "'" + StrColFolder + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
	
	'==============================DSI_FinishInstall_QuestSQLOptimizerForDB2zOS========================================
	Sub Update_DSI_FinishInstall_QuestSQLOptimizerForDB2zOS(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100	
		else
			select case UCase(StrProduct)
				case "SQLOPTIMIZERFORDB2ZOS_X86_EN"
					StrProduct="Dell% SQL Optimizer for IBM% DB2% z_OS_"
				case "SQLOPTIMIZERFORDB2ZOS_X64_EN"
					StrProduct="Dell% SQL Optimizer for IBM% DB2% z_OS_"
			end select
		end if
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_QuestSQLOptimizerForDB2zOS set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		
		'Update I_SubFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_SubFolder from DSI.dbo.DSI_FinishInstall_QuestSQLOptimizerForDB2zOS where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_SubFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_QuestSQLOptimizerForDB2zOS set  I_SubFolder =" + "'" + StrColFolder + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub

	'==============================DSI_FinishInstall_BMF========================================
	Sub Update_DSI_FinishInstall_BMF(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer,StrColDisplay
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case StrProduct
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="Benchmark Factory_ for Databases"
				case "BENCHMARKFACTORY_X64_EN"
					StrProduct="Benchmark Factory_ for Databases"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="Benchmark Factory_ for Databases Trial"
				case "BENCHMARKFACTORY_TRIAL_X64_EN"
					StrProduct="Benchmark Factory_ for Databases Trial"
			end select	
		end if
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_DB2_BMF set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_DB2_BMF where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(2)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_DB2_BMF set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		
		'Update I_DisplayVersion Column Record
		Query		= 	"Select I_DisplayVersion from DSI.dbo.DSI_DB2_BMF where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColDisplay	=	Rec.Fields("I_DisplayVersion").Value
			Rec.MoveNext
		Wend
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(3)
		if InStr(StrColDisplay,"32-bit") >= 3 then
			StrColDisplay	=	StrMainVer(0) + "." + StrMainVer(1) + " (32-bit)" + "." + StrMainVer(3)
		elseif InStr(StrColDisplay,"64-bit") >= 3  then
			StrColDisplay	=	StrMainVer(0) + "." + StrMainVer(1) + " (64-bit)" + "." + StrMainVer(3)
		else
			StrColDisplay 	= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(3)
		end if
		
		Conn.Execute "Update DSI.dbo.DSI_DB2_BMF set  I_DisplayVersion =" + "'" + StrColDisplay + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub

	'==============================DSI_FinishInstall_SpotlightonIBMDB2========================================
	Sub Update_DSI_FinishInstall_SpotlightonIBMDB2(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case StrProduct
				case "SPOTLIGHTONDB2_X86_EN"
					StrProduct	=	"Spotlight_ on IBM_ DB2_ LUW"
				case "SPOTLIGHTONDB2_X64_EN"
					StrProduct	=	"Spotlight_ on IBM_ DB2_ LUW"
			end select	
		end if
		
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_SpotlightonIBMDB2 set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		
		'Update I_SubFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_SubFolder from DSI.dbo.DSI_FinishInstall_SpotlightonIBMDB2 where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_SubFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_SpotlightonIBMDB2 set  I_SubFolder =" + "'" + StrColFolder + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub

	'==============================DSI_FinishInstall_ToadDataModeler========================================
	Sub Update_DSI_FinishInstall_ToadDataModeler(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100	
		else
			select case UCase(StrProduct)
				case "TOADDATAMODELER_X86_EN"
					StrProduct="TOAD% DATA MODELER 32-BIT"
                                case "TOADDATAMODELER_X64_EN"
					StrProduct="TOAD% DATA MODELER 64-BIT"
			end select
		end if
		
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_DB2_ToadDataModeler set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_DB2_ToadDataModeler where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_DB2_ToadDataModeler set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
	
	'==============================DSI_ProductSelectionPage_VerifyProductDetails========================================
	Sub Update_DSI_ProductSelectionPage_VerifyProductDetails(ByVal StrProduct,ByVal StrVersion)
		
		on error resume next
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "TOADFORDB2_X86_EN"
					StrProduct="TOAD_ FOR IBM_ DB2_"
				case "TOADFORDB2_TRIAL_X86_EN"
					StrProduct="TOAD_ FOR IBM_ DB2_ TRIAL"
				case "TOADDATAMODELER_X86_EN"
					StrProduct="TOAD% DATA MODELER 32-BIT"
                                case "TOADDATAMODELER_X64_EN"
					StrProduct="TOAD% DATA MODELER 64-BIT"
				case "SPOTLIGHTONDB2_X86_EN"
					StrProduct="SPOTLIGHT_ ON IBM_ DB2_ LUW"
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="BENCHMARK FACTORY_ FOR DATABASES"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="BENCHMARK FACTORY_ FOR DATABASES TRIAL"
				case "SQLOPTIMIZERFORDB2LUW_X86_EN"
					StrProduct="Dell_ SQL OPTIMIZER FOR IBM_ DB2% LUW"
				case "SQLOPTIMIZERFORDB2ZOS_X86_EN"
					StrProduct="Dell_ SQL OPTIMIZER FOR IBM_ DB2% Z_OS_"
				case else
					StrProduct="Null"
			end select
		end if
		
		Conn.Execute "Update DSI.dbo.DSI_ProductSelectionPage_VerifyProductDetails set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
	
	'==============================DSI_FinishInstall_VerifyRegistry========================================
	Sub Update_DSI_FinishInstall_VerifyRegistry(ByVal StrProduct,ByVal StrVersion)
		
		Dim StrColName,StrMainVer,Query,StrVer
		Dim Matches,match,RetStr
		
		'On error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "TOADFORDB2_X86_EN"
					StrProduct="TOAD_ FOR IBM_ DB2_"
				case "TOADFORDB2_TRIAL_X86_EN"
					StrProduct="TOAD_ FOR IBM_ DB2_ TRIAL"
				case "TOADDATAMODELER_X86_EN"
					StrProduct="TOAD% DATA MODELER 32-BIT"
                                case "TOADDATAMODELER_X64_EN"
					StrProduct="TOAD% DATA MODELER 64-BIT"
				case "SPOTLIGHTONDB2_X86_EN"
					StrProduct="SPOTLIGHT_ ON IBM_ DB2_ LUW"
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="BENCHMARK FACTORY_ FOR DATABASES"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="BENCHMARK FACTORY_ FOR DATABASES TRIAL"
				case "SQLOPTIMIZERFORDB2LUW_X86_EN"
					StrProduct="Dell_ SQL OPTIMIZER FOR IBM_ DB2% LUW"
				case "SQLOPTIMIZERFORDB2ZOS_X86_EN"
					StrProduct="Dell_ SQL OPTIMIZER FOR IBM_ DB2% Z_OS_"
				case else
					StrProduct="Null"
			end select
		end if
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_VerifyRegistry set  I_ProductVersion =" + "'" + StrVersion + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_InstallerDisplayProductName) like '" + StrProduct + "'"
		
		'Update I_ProductName Column
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_ProductName from DSI.dbo.DSI_FinishInstall_VerifyRegistry where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_InstallerDisplayProductName) like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColName=Rec.Fields("I_ProductName").Value
			Rec.MoveNext
		Wend
		StrMainVer 	= 	Split(StrVersion,".")
		if InStr(StrColName,"Benchmark Factory") >=	1 then
			StrVer 	= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(2)
		else
			StrVer 	= 	StrMainVer(0) + "." + StrMainVer(1)
		end if
		
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		Set Matches		=	RegEx.Execute(StrColName)
		For each match in matches
			RetStr		=	Match.Value
		Next
		if RetStr <> "" then
			StrColName 	= 	regEx.Replace(StrColName,StrVer)
			Conn.Execute "Update DSI.dbo.DSI_FinishInstall_VerifyRegistry set  I_ProductName =" + "'" + StrColName + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_InstallerDisplayProductName) like '" + StrProduct + "'"
		end if
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
	
	'==============================DSI_SilentInstallMsiBuild========================================
	Sub Update_SilentInstallMsiBuild(ByVal StrProduct,ByVal StrVersion)
		
		Dim StrColName,Query,StrVer
		Dim Matches,match,RetStr
		
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "TOADFORDB2_X86_EN"
					StrProduct="TOAD_ FOR IBM_ DB2_"
				case "TOADFORDB2_TRIAL_X86_EN"
					StrProduct="TOAD_ FOR IBM_ DB2_ TRIAL"
				case "TOADDATAMODELER_X86_EN"
					StrProduct="TOAD% DATA MODELER 32-BIT"
                                case "TOADDATAMODELER_X64_EN"
					StrProduct="TOAD% DATA MODELER 64-BIT"
				case "SPOTLIGHTONDB2_X86_EN"
					StrProduct="SPOTLIGHT_ ON IBM_ DB2_ LUW"
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="BENCHMARK FACTORY_ FOR DATABASES"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="BENCHMARK FACTORY_ FOR DATABASES TRIAL"
				case "SQLOPTIMIZERFORDB2LUW_X86_EN"
					StrProduct="Dell_ SQL OPTIMIZER FOR IBM_ DB2% LUW"
				case "SQLOPTIMIZERFORDB2ZOS_X86_EN"
					StrProduct="Dell_ SQL OPTIMIZER FOR IBM_ DB2% Z_OS_"
				case else
					StrProduct="Null"
			end select
		end if
		
		'Update I_FilePath Column
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_FilePath from DSI.dbo.DB2_SilentInstallMsiBuild where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColName=Rec.Fields("I_FilePath").Value
			Rec.MoveNext
		Wend
		
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		Set Matches		=	RegEx.Execute(StrColName)
		For each match in matches
			RetStr		=	Match.Value
		Next
		if RetStr <> "" then
			StrColName 	= 	regEx.Replace(StrColName,StrVersion)
			Conn.Execute "Update DSI.dbo.DB2_SilentInstallMsiBuild set  I_FilePath =" + "'" + StrColName + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		end if
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
        '==============================DSI_ValidateShortcutAndKeyFile========================================
	Sub Update_DSI_DB2_ValidateShortcutAndKeyFile(ByVal StrProduct,ByVal StrVersion)
		
		Dim StrColName,StrMainVer,Query,StrVer
		Dim Matches,match,RetStr
		
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "TOADFORDB2_X86_EN"
					StrProduct="TOAD FOR IBM%"
				case "TOADFORDB2_TRIAL_X86_EN"
					StrProduct="TOAD FOR IBM%"
				case "TOADDATAMODELER_X86_EN"
					StrProduct="TOAD DATA MODELER%"
                                case "TOADDATAMODELER_X64_EN"
					StrProduct="TOAD DATA MODELER%"
				case "SPOTLIGHTONDB2_X86_EN"
					StrProduct="SPOTLIGHT ON IBM DB2 LUW%"
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="BENCHMARK FACTORY FOR%"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="BENCHMARK FACTORY FOR%"
				case else
					StrProduct="Null"
			end select
		end if

		
		'Update I_ProductName Column
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_ProductName from DSI.dbo.DSI_DB2_ValidateShortcutAndKeyFile where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColName=Rec.Fields("I_ProductName").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")

		if InStr(StrColName,"Spotlight on IBM DB2 LUW") >=	1 then
			StrVer 	= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(2)
		elseif InStr(StrColName,"Benchmark Factory") >=	1 then
                        StrVer 	= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(2)
                else
			StrVer 	= 	StrMainVer(0) + "." + StrMainVer(1)
		end if
		
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		Set Matches		=	RegEx.Execute(StrColName)
		For each match in matches
			RetStr		=	Match.Value
		Next
		if RetStr <> "" then
			StrColName 	= 	regEx.Replace(StrColName,StrVer)
			Conn.Execute "Update DSI.dbo.DSI_DB2_ValidateShortcutAndKeyFile set  I_ProductName =" + "'" + StrColName + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		end if
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub

End Class

Class UpdateSQLServerSuite

	'==============================DSI_SQLServer _FinishInstall_BMF========================================
	Sub Update_DSI_SQLServer_FinishInstall_BMF(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case StrProduct
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="Benchmark Factory_ for Databases"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="Benchmark Factory_ for Databases Trial"
			end select	
		end if
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_SQLServer_FinishInstall_BMF set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_SQLServer_FinishInstall_BMF where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(2)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_SQLServer_FinishInstall_BMF set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		
		'Update I_DisplayVersion Column Record
		Query		= 	"Select I_DisplayVersion from DSI.dbo.DSI_SQLServer_FinishInstall_BMF where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColDisplay	=	Rec.Fields("I_DisplayVersion").Value
			Rec.MoveNext
		Wend
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(3)
		if InStr(StrColDisplay,"32-bit") >= 3 then
			StrColDisplay	=	StrMainVer(0) + "." + StrMainVer(1) + " (32-bit)" + "." + StrMainVer(3)
		elseif InStr(StrColDisplay,"64-bit") >= 3  then
			StrColDisplay	=	StrMainVer(0) + "." + StrMainVer(1) + " (64-bit)" + "." + StrMainVer(3)
		else
			StrColDisplay 	= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(3)
		end if
		
		Conn.Execute "Update DSI.dbo.DSI_SQLServer_FinishInstall_BMF set  I_DisplayVersion =" + "'" + StrColDisplay + "'" + " where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub


	'==============================DSI_SQLServer_FinishInstall_ToadDataModeler========================================
	Sub Update_DSI_SQLServer_FinishInstall_ToadDataModeler(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100	
		else
			select case UCase(StrProduct)
				case "TOADDATAMODELER_X86_EN"
					StrProduct="TOAD% DATA MODELER 32-BIT"
                                case "TOADDATAMODELER_X64_EN"
					StrProduct="TOAD% DATA MODELER 64-BIT"
			end select
		end if
		
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_SQLServer_FinishInstall_ToadDataModeler set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_SQLServer_FinishInstall_ToadDataModeler where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_SQLServer_FinishInstall_ToadDataModeler set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like '" + StrProduct + "'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
	
	'==============================DSI_FinishInstall_SoSSE========================================
	Sub Update_DSI_FinishInstall_SoSSE(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		end if
		
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_SoSSE set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and Upper(I_ProductName) like 'SPOTLIGHT_ ON SQL SERVER STANDARD'"
		
		'Update I_SubFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_FinishInstall_SoSSE where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like 'SPOTLIGHT_ ON SQL SERVER STANDARD'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_SoSSE set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like 'SPOTLIGHT_ ON SQL SERVER STANDARD'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
	
	'==============================DSI_FinishInstall_QSOSS========================================
	Sub Update_DSI_FinishInstall_QSOSS(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100	
		else
			select case UCase(StrProduct)
				case "SQLOPTIMIZERFORSQLSERVER_X86_EN"
					StrProduct="DELL% SQL OPTIMIZER FOR SQL SERVER"
				case "SQLOPTIMIZERFORSQLSERVER_TRIAL_X86_EN"
					StrProduct="DELL% SQL OPTIMIZER FOR SQL SERVER TRIAL"
			end select
		end if
		
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_QSOSS set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		
		'Update I_SubFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_FinishInstall_QSOSS where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_QSOSS set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
	
	'==============================DSI_FinishInstall_ToadforSQLServer========================================
	Sub Update_DSI_FinishInstall_ToadforSQLServer(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "TOADFORSQLSERVER_X86_EN"
					StrProduct="TOAD_ FOR SQL SERVER"
				case "TOADFORSQLSERVER_TRIAL_X86_EN"
					StrProduct="TOAD_ FOR SQL SERVER TRIAL"
			end select	
		end if
		
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_ToadforSQLServer set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct +"'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_FinishInstall_ToadforSQLServer where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and Upper(I_ProductName) like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_ToadforSQLServer set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" +  StrProduct + "'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
	
	'==============================DSI_SQLServer_VerifyRegistry========================================
	Sub Update_DSI_SQLServer_VerifyRegistry(ByVal StrProduct,ByVal StrVersion)
		
		Dim StrColName,StrMainVer,Query,StrVer
		Dim Matches,match,RetStr
		
		On error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "TOADFORSQLSERVER_X86_EN"
					StrProduct="TOAD_ FOR SQL SERVER"
				case "TOADFORSQLSERVER_TRIAL_X86_EN"
					StrProduct="TOAD_ FOR SQL SERVER TRIAL"
				case "TOADDATAMODELER_X86_EN"
					StrProduct="TOAD% DATA MODELER 32-BIT"
                                case "TOADDATAMODELER_X64_EN"
					StrProduct="TOAD% DATA MODELER 64-BIT"
				case "SPOTLIGHTONSQLSERVER_STANDARD_X86_EN"
					StrProduct="SPOTLIGHT_ ON SQL SERVER STANDARD"
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="BENCHMARK FACTORY_ FOR DATABASES"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="BENCHMARK FACTORY_ FOR DATABASES TRIAL"
				case "SQLOPTIMIZERFORSQLSERVER_X86_EN"
					StrProduct="DELL_ SQL OPTIMIZER FOR SQL SERVER"
				case "SQLOPTIMIZERFORSQLSERVER_TRIAL_X86_EN"
					StrProduct="DELL_ SQL OPTIMIZER FOR SQL SERVER TRIAL"
				case else
					StrProduct="Null"
			end select
		end if
		'Update I_Version Column
		'wscript.echo("The product name [" + StrProduct + "], the version [" + StrVersion + "]")
		Conn.Execute "Update DSI.dbo.DSI_SQLServer_VerifyRegistry set  I_ProductVersion =" + "'" + StrVersion + "'" + " where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_InstallerDisplayProductName) like '" + StrProduct + "'"
		
		'Update I_ProductName Column
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_ProductName from DSI.dbo.DSI_SQLServer_VerifyRegistry where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_InstallerDisplayProductName) like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColName=Rec.Fields("I_ProductName").Value
			Rec.MoveNext
		Wend
		StrMainVer 	= 	Split(StrVersion,".")
		if InStr(StrColName,"Benchmark Factory") >=	1 then
			StrVer 	= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(2)
		elseif InStr(StrColName,"Dell SQL Optimizer for SQL Server") >=	1 then
			StrVer 	= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(2)
		else
			StrVer 	= 	StrMainVer(0) + "." + StrMainVer(1)
		end if
		
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		Set Matches		=	RegEx.Execute(StrColName)
		For each match in matches
			RetStr		=	Match.Value
		Next
		if RetStr <> "" then
			StrColName 	= 	regEx.Replace(StrColName,StrVer)
			Conn.Execute "Update DSI.dbo.DSI_SQLServer_VerifyRegistry set  I_ProductName =" + "'" + StrColName + "'" + " where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_InstallerDisplayProductName) like '" + StrProduct + "'"
		end if
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
	
	'==============================DSI_SQLServer_VerifyProductDetails========================================
	Sub Update_DSI_SQLServer_VerifyProductDetails(ByVal StrProduct,ByVal StrVersion)
		
		on error resume next
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "TOADFORSQLSERVER_X86_EN"
					StrProduct="TOAD_ FOR SQL SERVER"
				case "TOADFORSQLSERVER_TRIAL_X86_EN"
					StrProduct="TOAD_ FOR SQL SERVER TRIAL"
				case "TOADDATAMODELER_X86_EN"
					StrProduct="TOAD% DATA MODELER 32-BIT"
                                case "TOADDATAMODELER_X64_EN"
					StrProduct="TOAD% DATA MODELER 64-BIT"
				case "SPOTLIGHTONSQLSERVER_STANDARD_X86_EN"
					StrProduct="SPOTLIGHT_ ON SQL SERVER STANDARD"
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="BENCHMARK FACTORY_ FOR DATABASES"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="BENCHMARK FACTORY_ FOR DATABASES TRIAL"
				case "SQLOPTIMIZERFORSQLSERVER_X86_EN"
					StrProduct="Dell_ SQL OPTIMIZER FOR SQL SERVER"
				case "SQLOPTIMIZERFORSQLSERVER_TRIAL_X86_EN"
					StrProduct="Dell_ SQL OPTIMIZER FOR SQL SERVER TRIAL"
				case else
					StrProduct="Null"
			end select
		end if
		
		Conn.Execute "Update DSI.dbo.DSI_SQLServer_VerifyProductDetail set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
	
	'==============================SQLServer_SilentInstallMsiBuild========================================
	Sub Update_SQLServer_SilentInstallMsiBuild(ByVal StrProduct,ByVal StrVersion)
		
		Dim StrColName,Query,StrVer
		Dim Matches,match,RetStr
		
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "TOADFORSQLSERVER_X86_EN"
					StrProduct="TOAD_ FOR SQL SERVER"
				case "TOADFORSQLSERVER_TRIAL_X86_EN"
					StrProduct="TOAD_ FOR SQL SERVER TRIAL"
				case "TOADDATAMODELER_X86_EN"
					StrProduct="TOAD% DATA MODELER 32-BIT"
                                case "TOADDATAMODELER_X64_EN"
					StrProduct="TOAD% DATA MODELER 64-BIT"
				case "SPOTLIGHTONSQLSERVER_STANDARD_X86_EN"
					StrProduct="SPOTLIGHT_ ON SQL SERVER STANDARD"
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="BENCHMARK FACTORY_ FOR DATABASES"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="BENCHMARK FACTORY_ FOR DATABASES TRIAL"
				case "SQLOPTIMIZERFORSQLSERVER_X86_EN"
					StrProduct="Dell_ SQL OPTIMIZER FOR SQL SERVER"
				case "SQLOPTIMIZERFORSQLSERVER_TRIAL_X86_EN"
					StrProduct="Dell_ SQL OPTIMIZER FOR SQL SERVER TRIAL"
				case else
					StrProduct="Null"
			end select
		end if
		
		'Update I_FilePath Column
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_FilePath from DSI.dbo.SQLServer_SilentInstallMsiBuild where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColName=Rec.Fields("I_FilePath").Value
			Rec.MoveNext
		Wend
		
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		Set Matches		=	RegEx.Execute(StrColName)
		For each match in matches
			RetStr		=	Match.Value
		Next
		if RetStr <> "" then
			StrColName 	= 	regEx.Replace(StrColName,StrVersion)
			Conn.Execute "Update DSI.dbo.SQLServer_SilentInstallMsiBuild set  I_FilePath =" + "'" + StrColName + "'" + " where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		end if
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
'==============================DSI_ValidateShortcutAndKeyFile========================================
	Sub Update_DSI_SQLServer_ValidateShortcutAndKeyFile(ByVal StrProduct,ByVal StrVersion)
		
		Dim StrColName,StrMainVer,Query,StrVer
		Dim Matches,match,RetStr
		
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "TOADFORSQLSERVER_X86_EN"
					StrProduct="TOAD FOR SQL SERVER%"
				case "TOADFORSQLSERVER_TRIAL_X86_EN"
					StrProduct="TOAD FOR SQL SERVER%"
				case "TOADDATAMODELER_X86_EN"
					StrProduct="TOAD DATA MODELER%"
                                case "TOADDATAMODELER_X64_EN"
					StrProduct="TOAD DATA MODELER%"
				case "SPOTLIGHTONSQLSERVER_STANDARD_X86_EN"
					StrProduct="SPOTLIGHT ON SQL SERVER%"
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="BENCHMARK FACTORY%"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="BENCHMARK FACTORY%"
				case else
					StrProduct="Null"
			end select
		end if

		
		'Update I_ProductName Column
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_ProductName from DSI.dbo.DSI_SQLServer_ValidateShortcutAndKeyFile where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColName=Rec.Fields("I_ProductName").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")

		if InStr(StrColName,"Benchmark Factory") >=	1 then
                        StrVer 	= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(2)
                else
			StrVer 	= 	StrMainVer(0) + "." + StrMainVer(1)
		end if
		
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		Set Matches		=	RegEx.Execute(StrColName)
		For each match in matches
			RetStr		=	Match.Value
		Next
		if RetStr <> "" then
			StrColName 	= 	regEx.Replace(StrColName,StrVer)
			Conn.Execute "Update DSI.dbo.DSI_SQLServer_ValidateShortcutAndKeyFile set  I_ProductName =" + "'" + StrColName + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		end if
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub

End Class

Class UpdateSQLNavigatorSuite
'==============================DSI_ProductSelectionPage_VerifyProductDetail========================================
	Sub Update_DSI_SQLNavigator_VerifyProductDetail(ByVal StrProduct,ByVal StrVersion)
		
		on error resume next
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "SQLNavigator_X64_EN"
					StrProduct="SQL Navigator 64-BIT"
				case "SQLNavigator_X86_EN"
					StrProduct="SQL Navigator 32-BIT"
				case "SQLNavigator_TRIAL_X86_EN"
					StrProduct="SQL Navigator 32-BIT TRIAL"
				case "SQLNavigator_TRIAL_X64_EN"
					StrProduct="SQL Navigator 64-BIT TRIAL"
				case "CODETESTERORACLE_X86_EN"
					StrProduct="DELL% CODE TESTER FOR ORACLE"
				case "BENCHMARKFACTORY_X64_EN"
					StrProduct="BENCHMARK FACTORY% 64-BIT"
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="BENCHMARK FACTORY% 32-BIT"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="BENCHMARK FACTORY% 32-BIT Trial"
				case "BENCHMARKFACTORY_TRIAL_X64_EN"
					StrProduct="BENCHMARK FACTORY% 64-BIT Trial"
				case "SQLOPTIMIZERFORORACLE_X64_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 64-BIT"
				case "SQLOPTIMIZERFORORACLE_X86_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 32-BIT"
				case "SQLOPTIMIZERFORORACLE_TRIAL_X86_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 32-BIT TRIAL"
				case "SQLOPTIMIZERFORORACLE_TRIAL_X64_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 64-BIT TRIAL"
				case else
					StrProduct="Null"
			end select
		end if
		
		Conn.Execute "Update DSI.dbo.DSI_SQLNavigator_VerifyProductDetail set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 5 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
 '==============================DSI_FinishInstall_SQLNavigator========================================
	Sub Update_DSI_FinishInstall_SQLNavigator(ByVal StrProduct,ByVal StrVersion)
		
		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "SQLNavigator_X64_EN"
					StrProduct="64-bit"
				case "SQLNavigator_X86_EN"
					StrProduct="32-bit"
				case "SQLNavigator_TRIAL_X86_EN"
					StrProduct="32-bit Trial"
				case "SQLNavigator_TRIAL_X64_EN"
					StrProduct="64-bit Trial"

			end select	
		end if
		'Update I_Version Column Record
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_SQLNavigator set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 5 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'SQL Navigator%" + StrProduct +"'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_FinishInstall_SQLNavigator where Projectid = 5 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'SQL Navigator%" + StrProduct +"'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_SQLNavigator set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 5 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'SQL Navigator%" + StrProduct +"'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
 '==============================DSI_FinishInstall_QuestCodeTester========================================
	Sub Update_DSI_FinishInstall_QuestCodeTester(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		end if
		
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_QuestCodeTester set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 5 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like 'DELL_ CODE TESTER FOR ORACLE'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_FinishInstall_QuestCodeTester where Projectid = 5 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'DELL_ CODE TESTER FOR ORACLE'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_QuestCodeTester set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 5 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'DELL_ CODE TESTER FOR ORACLE'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
 '==============================DSI_FinshInstall_OptimizerforOracle========================================
	Sub Update_DSI_FinshInstall_OptimizerforOracle(ByVal StrProduct,ByVal StrVersion)
		
		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
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
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_FinshInstall_OptimizerforOracle set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 5 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Dell% SQL Optimizer for Oracle%" + StrProduct +"'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_FinshInstall_OptimizerforOracle where Projectid = 5 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Dell% SQL Optimizer for Oracle%" + StrProduct +"'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_FinshInstall_OptimizerforOracle set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 5 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Dell% SQL Optimizer for Oracle%" + StrProduct +"'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
 '==============================DSI_FinishInstall_BMF========================================
	Sub Update_DSI_FinishInstall_BMF(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer,StrColDisplay
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case StrProduct
				case "BENCHMARKFACTORY_X64_EN"
					StrProduct="BENCHMARK FACTORY% 64-BIT"
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="BENCHMARK FACTORY% 32-BIT"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="BENCHMARK FACTORY% 32-BIT TRIAL"
				case "BENCHMARKFACTORY_TRIAL_X64_EN"
					StrProduct="BENCHMARK FACTORY% 64-BIT TRIAL"
			end select	
		end if
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_BMF set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 5 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_FinishInstall_BMF where Projectid = 5 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColFolder=Rec.Fields("I_InstallFolder").Value
			Rec.MoveNext
		Wend
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(2)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_BMF set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 5 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		
		'Update I_DisplayVersion Column Record
		Query		= 	"Select I_DisplayVersion from DSI.dbo.DSI_FinishInstall_BMF where Projectid = 5 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColDisplay	=	Rec.Fields("I_DisplayVersion").Value
			Rec.MoveNext
		Wend
		StrMainVer 	= 	Split(StrVersion,".")
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(3)
		if InStr(StrColDisplay,"32-bit") >= 3 then
			StrColDisplay	=	StrMainVer(0) + "." + StrMainVer(1) + " (32-bit)" + "." + StrMainVer(3)
		elseif InStr(StrColDisplay,"64-bit") >= 3  then
			StrColDisplay	=	StrMainVer(0) + "." + StrMainVer(1) + " (64-bit)" + "." + StrMainVer(3)
		else
			StrColDisplay 	= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(3)
		end if
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_BMF set  I_DisplayVersion =" + "'" + StrColDisplay + "'" + " where Projectid = 5 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
'==============================DSI_FinishInstall_VerifyRegistry========================================
	Sub Update_DSI_FinishInstall_VerifyRegistry(ByVal StrProduct,ByVal StrVersion)
		
		Dim StrColName,StrMainVer,Query,StrVer
		Dim Matches,match,RetStr
		
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "SQLNavigator_X64_EN"
					StrProduct="SQL Navigator 64-BIT"
				case "SQLNavigator_X86_EN"
					StrProduct="SQL Navigator 32-BIT"
				case "SQLNavigator_TRIAL_X86_EN"
					StrProduct="SQL Navigator 32-BIT TRIAL"
				case "SQLNavigator_TRIAL_X64_EN"
					StrProduct="SQL Navigator 64-BIT TRIAL"
				case "CODETESTERORACLE_X86_EN"
					StrProduct="DELL% CODE TESTER FOR ORACLE"
				case "BENCHMARKFACTORY_X64_EN"
					StrProduct="BENCHMARK FACTORY% 64-BIT"
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="BENCHMARK FACTORY% 32-BIT"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="BENCHMARK FACTORY% 32-BIT Trial"
				case "BENCHMARKFACTORY_TRIAL_X64_EN"
					StrProduct="BENCHMARK FACTORY% 64-BIT Trial"
				case "SQLOPTIMIZERFORORACLE_X64_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 64-BIT"
				case "SQLOPTIMIZERFORORACLE_X86_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 32-BIT"
				case "SQLOPTIMIZERFORORACLE_TRIAL_X86_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 32-BIT TRIAL"
				case "SQLOPTIMIZERFORORACLE_TRIAL_X64_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 64-BIT TRIAL"
				case else
					StrProduct="Null"
			end select
		end if
		'Update I_ProductVersion Column
		Conn.Execute "Update DSI.dbo.DSI_SQLNavigator_VerifyRegistry set  I_ProductVersion =" + "'" + StrVersion + "'" + " where Projectid = 5 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_InstallerDisplayProductName) like '" + StrProduct + "'"
		
		'Update I_ProductName Column
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_ProductName from DSI.dbo.DSI_SQLNavigator_VerifyRegistry where Projectid = 5 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_InstallerDisplayProductName) like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColName=Rec.Fields("I_ProductName").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")
		
		if InStr(StrColName,"Benchmark Factory") >=	1 then
			StrVer 	= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(2)
		else
			StrVer 	= 	StrMainVer(0) + "." + StrMainVer(1)
		end if
		
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		Set Matches		=	RegEx.Execute(StrColName)
		For each match in matches
			RetStr		=	Match.Value
		Next
		if RetStr <> "" then
			StrColName 	= 	regEx.Replace(StrColName,StrVer)
			Conn.Execute "Update DSI.dbo.DSI_SQLNavigator_VerifyRegistry set  I_ProductName =" + "'" + StrColName + "'" + " where Projectid = 5 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_InstallerDisplayProductName) like '" + StrProduct + "'"
		end if
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub

	'==============================DSI_ValidateShortcutAndKeyFile========================================
	Sub Update_DSI_ValidateShortcutAndKeyFile(ByVal StrProduct,ByVal StrVersion)
		
		Dim StrColName,StrMainVer,Query,StrVer
		Dim Matches,match,RetStr
		
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "SQLNavigator_X64_EN"
					StrProduct="SQL Navigator [1-9].[0-9]"
				case "SQLNavigator_X86_EN"
					StrProduct="SQL Navigator [1-9].[0-9]"
				case "SQLNavigator_TRIAL_X86_EN"
					StrProduct="SQL Navigator [1-9].[0-9] TRIAL"
				case "SQLNavigator_TRIAL_X64_EN"
					StrProduct="SQL Navigator [1-9].[0-9] TRIAL"
				case "CODETESTERORACLE_X86_EN"
					StrProduct="DELL CODE TESTER FOR ORACLE%"
				case "SPOTLIGHTONORACLE_X64_MULTILANG"
					StrProduct="SPOTLIGHT ON ORACLE%"
				case "SPOTLIGHTONORACLE_X86_MULTILANG"
					StrProduct="SPOTLIGHT ON ORACLE%"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="BENCHMARK FACTORY TRIAL [1-9].[0-9].[0-9]"
				case "BENCHMARKFACTORY_TRIAL_X64_EN"
					StrProduct="BENCHMARK FACTORY TRIAL [1-9].[0-9].[0-9](64-bit)"
                                case "BENCHMARKFACTORY_X64_EN"					
                                        StrProduct="BENCHMARK FACTORY [1-9].[0-9].[0-9](64-bit)"				
                                case "BENCHMARKFACTORY_X86_EN"
                                     StrProduct="BENCHMARK FACTORY [1-9].[0-9].[0-9]"
				case "SQLOPTIMIZERFORORACLE_X64_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 64-BIT"
				case else
					StrProduct="Null"
			end select
		end if

		
		'Update I_ProductName Column
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_ProductName from DSI.dbo.DSI_SQLNavigator_ValidateShortcutAndKeyFile where Projectid = 5 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColName=Rec.Fields("I_ProductName").Value
			Rec.MoveNext
		Wend
		
		StrMainVer 	= 	Split(StrVersion,".")

		if (InStr(StrColName,"Benchmark Factory") >=	1) or (InStr(StrColName,"Benchmark Factory Trial") >=	1) then
			StrVer 	= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(2)
		else
			StrVer 	= 	StrMainVer(0) + "." + StrMainVer(1)
		end if
		
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		Set Matches		=	RegEx.Execute(StrColName)
		For each match in matches
			RetStr		=	Match.Value
		Next
		if RetStr <> "" then
			StrColName 	= 	regEx.Replace(StrColName,StrVer)
			Conn.Execute "Update DSI.dbo.DSI_SQLNavigator_ValidateShortcutAndKeyFile set  I_ProductName =" + "'" + StrColName + "'" + " where Projectid = 5 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		end if
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
        
'==============================DSI_SilentInstallMsiBuild========================================
	Sub Update_SilentInstallMsiBuild(ByVal StrProduct,ByVal StrVersion)
		
		Dim StrColName,Query,StrVer
		Dim Matches,match,RetStr
		
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "SQLNavigator_X64_EN"
					StrProduct="SQL Navigator 64-BIT"
				case "SQLNavigator_X86_EN"
					StrProduct="SQL Navigator 32-BIT"
				case "SQLNavigator_TRIAL_X86_EN"
					StrProduct="SQL Navigator 32-BIT TRIAL"
				case "SQLNavigator_TRIAL_X64_EN"
					StrProduct="SQL Navigator 64-BIT TRIAL"
				case "CODETESTERORACLE_X86_EN"
					StrProduct="DELL% CODE TESTER FOR ORACLE"
				case "BENCHMARKFACTORY_X64_EN"
					StrProduct="BENCHMARK FACTORY% 64-BIT"
				case "BENCHMARKFACTORY_X86_EN"
					StrProduct="BENCHMARK FACTORY% 32-BIT"
				case "BENCHMARKFACTORY_TRIAL_X86_EN"
					StrProduct="BENCHMARK FACTORY% 32-BIT Trial"
				case "BENCHMARKFACTORY_TRIAL_X64_EN"
					StrProduct="BENCHMARK FACTORY% 64-BIT Trial"
				case "SQLOPTIMIZERFORORACLE_X64_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 64-BIT"
				case "SQLOPTIMIZERFORORACLE_X86_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 32-BIT"
				case "SQLOPTIMIZERFORORACLE_TRIAL_X86_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 32-BIT TRIAL"
				case "SQLOPTIMIZERFORORACLE_TRIAL_X64_MULTILANG"
					StrProduct="DELL% SQL OPTIMIZER FOR ORACLE 64-BIT TRIAL"
				case else
					StrProduct="Null"
			end select
		end if
		
		'Update I_FilePath Column
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_FilePath from DSI.dbo.DSI_SQLNavigator_SilentInstallMsiBuild where Projectid = 5 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		Set Rec		=	Conn.Execute(Query)
		While not Rec.EOF
			StrColName=Rec.Fields("I_FilePath").Value
			Rec.MoveNext
		Wend
		
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		Set Matches		=	RegEx.Execute(StrColName)
		For each match in matches
			RetStr		=	Match.Value
		Next
		if RetStr <> "" then
			StrColName 	= 	regEx.Replace(StrColName,StrVersion)
			Conn.Execute "Update DSI.dbo.DSI_SQLNavigator_SilentInstallMsiBuild set  I_FilePath =" + "'" + StrColName + "'" + " where Projectid = 5 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		end if
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number <> 0 then
			Err.Clear
		end if

	End Sub
End Class
'================================================================================

Sub UpdateTestData()

        on error resume next 

	Dim XMLDoc,FSO,regEx
	Dim ErrorMsg,i
	Dim NewOracleSuite,NewSAPSuite,NewDB2Suite,NewSQLServerSuite,NewSQLNavigator
	Dim ProjectFile
	Dim RootNode,ProductNode,NodeName
	Dim productName,productversion,StrProduct,PreProduct
	Dim ParentGroup,groupowner,childgroup
	Dim Matches,match

	If WScript.Arguments.Count	=	2 then
		ProjectFile	=	Trim(WScript.Arguments(0))
		StrProduct	=	Trim(WScript.Arguments(1))
	Else
		wscript.Quit 400
	End If
	

	set XMLDoc		=	CreateObject("MSXML2.DOMDOCUMENT")

	XMLDoc.async	=	False

	XMLDoc.ValidateonParse=True
	'Open project file
	Set FSO	=	CreateObject("Scripting.FileSystemObject")
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
	
	Set Conn	=	CreateObject("ADODB.Connection")
	Conn.Open "Driver=SQL Server;Server=10.6.208.62;Database=DSI;uid=sa;pwd=Quest6848;"	
	
	set ProductNode	=	Productnode.item(0)
	set regEx		= 	New RegExp
	
	For i = 0 to RootNode.childnodes.length - 1
		
		NodeName	=	Productnode.childnodes.item(i).text
		NodeName	=	Split(NodeName,"=")
			
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		set Matches 	=	regEx.Execute(NodeName(1))
			
		for each match in Matches
			ProductVersion	=	match.value
		next
		ProductName	=	Mid(NodeName(0),1,Len(NodeName(0)) - 15)
		if ProductName <> "" and ProductVersion <> "" then
			PreProduct=Split(ProductName,"_")
			Select Case UCase(StrProduct)
				case UCase("ORACLE")
					Set NewOracleSuite	=	New UpdateOracleSuite
					'Update all finish installation data
					Select Case Trim(UCase(PreProduct(0)))
						case "TOADFORORACLE"
							Call NewOracleSuite.Update_DSI_FinishInstall_ToadforOracle(ProductName,ProductVersion)
						case "SQLOPTIMIZERFORORACLE"
							Call NewOracleSuite.Update_DSI_FinshInstall_OptimizerforOracle(ProductName,ProductVersion)
						case "BENCHMARKFACTORY"
							Call NewOracleSuite.Update_DSI_FinishInstall_BMF(ProductName,ProductVersion)
						case "SPOTLIGHTONORACLE"
							Call NewOracleSuite.Update_DSI_FinishInstall_SpotlightonOracle(ProductName,ProductVersion)
						case "TOADDATAMODELER"
							Call NewOracleSuite.Update_DSI_FinishInstall_ToadDataModeler(ProductName,ProductVersion)
						case "CODETESTERORACLE"
							Call NewOracleSuite.Update_DSI_FinishInstall_QuestCodeTester(ProductName,ProductVersion) 
						case "BACKUPREPORTER"
							Call NewOracleSuite.Update_DSI_FinishInstall_BackupReportForOracle(ProductName,ProductVersion) 
						case "TOADFORMYSQL"
							Call NewOracleSuite.Update_DSI_FinishInstall_ToadforMySQLFreeware(ProductName,ProductVersion) 
					End Select
					'Update Product Details table data
					Call NewOracleSuite.Update_DSI_ProductSelectionPage_VerifyProductDetail(ProductName,ProductVersion)
					'Update Verify Reistry table data
					Call NewOracleSuite.Update_DSI_FinishInstall_VerifyRegistry(ProductName,ProductVersion) 
					'Update Silent Install table data
					Call NewOracleSuite.Update_SilentInstallMsiBuild(ProductName,ProductVersion)
                                        'Update ShortCut table data
					Call NewOracleSuite.Update_DSI_ValidateShortcutAndKeyFile(ProductName,ProductVersion)

				case UCase("DB2")
					Set NewDB2Suite	=	New UpdateDB2Suite
					'Update all finish installation data
					Select Case Trim(UCase(PreProduct(0)))
						case "TOADFORDB2"
							Call NewDB2Suite.Update_DSI_FinishInstall_ToadforIBMDB2LUW(ProductName,ProductVersion)	
						case "SQLOPTIMIZERFORDB2LUW"
							Call NewDB2Suite.Update_DSI_FinishInstall_QuestSQLOptimizerforIBMDB2(ProductName,ProductVersion)
						case "SQLOPTIMIZERFORDB2ZOS"
							Call NewDB2Suite.Update_DSI_FinishInstall_QuestSQLOptimizerForDB2zOS(ProductName,ProductVersion)
						case "BENCHMARKFACTORY"
							Call NewDB2Suite.Update_DSI_FinishInstall_BMF(ProductName,ProductVersion)
						case "SPOTLIGHTONDB2"
							Call NewDB2Suite.Update_DSI_FinishInstall_SpotlightonIBMDB2(ProductName,ProductVersion)
						case "TOADDATAMODELER"
							Call NewDB2Suite.Update_DSI_FinishInstall_ToadDataModeler(ProductName,ProductVersion)
					End Select
					'Update Product Details table data
					Call NewDB2Suite.Update_DSI_ProductSelectionPage_VerifyProductDetails(ProductName,ProductVersion) 
					'Update Verify Reistry table data
					Call NewDB2Suite.Update_DSI_FinishInstall_VerifyRegistry(ProductName,ProductVersion)
					'Update Silent Install table data
					Call NewDB2Suite.Update_SilentInstallMsiBuild(ProductName,ProductVersion)
					'Update Shotcut table data
					Call NewDB2Suite.Update_DSI_DB2_ValidateShortcutAndKeyFile(ProductName,ProductVersion)
				case UCase("SAP")
					Set NewSAPSuite	=	New UpdateSAPSuite
					'Update all finish installation data
					Select Case Trim(UCase(PreProduct(0)))
						case "TOADFORSAP"
							Call NewSAPSuite.Update_DSI_FinishInstall_ToadforSybase(ProductName,ProductVersion) 
						case "SQLOPTIMIZERFORSAP"
							Call NewSAPSuite.Update_DSI_FinishInstall_QuestSQLOptimizerforSybase(ProductName,ProductVersion)
						case "BENCHMARKFACTORY"
							Call NewSAPSuite.Update_DSI_FinishInstall_BMF(ProductName,ProductVersion)
						case "SPOTLIGHTONSAP"
							Call NewSAPSuite.Update_DSI_FinishInstall_SpotlightonSybase(ProductName,ProductVersion)
						case "TOADDATAMODELER"
							Call NewSAPSuite.Update_DSI_FinishInstall_ToadDataModeler(ProductName,ProductVersion)
					End Select
					'Update Product Details table data
					Call NewSAPSuite.Update_DSI_ProductSelectionPage_VerifyProductDetails(ProductName,ProductVersion)
					'Update Verify Reistry table data
					Call NewSAPSuite.Update_DSI_FinishInstall_VerifyRegistry(ProductName,ProductVersion)
					'Update Silent Install table data
					Call NewSAPSuite.Update_SilentInstallMsiBuild(ProductName,ProductVersion)
                                        'Update Shortcut table data
					Call NewSAPSuite.Update_DSI_FinishInstall_ValidateShortcut(ProductName,ProductVersion)
				case UCase("SQLSERVER")
					Set NewSQLServerSuite	=	New UpdateSQLSERVERSuite
					'Update all finish installation data
					Select Case Trim(UCase(PreProduct(0)))
						case "TOADFORSQLSERVER"
							Call NewSQLServerSuite.Update_DSI_FinishInstall_ToadforSQLServer(ProductName,ProductVersion) 
						case "SQLOPTIMIZERFORSQLSERVER"
							Call NewSQLServerSuite.Update_DSI_FinishInstall_QSOSS(ProductName,ProductVersion)
						case "BENCHMARKFACTORY"
							Call NewSQLServerSuite.Update_DSI_SQLServer_FinishInstall_BMF(ProductName,ProductVersion)
						case "SPOTLIGHTONSQLSERVER"
							Call NewSQLServerSuite.Update_DSI_FinishInstall_SoSSE(ProductName,ProductVersion)
						case "TOADDATAMODELER"
							Call NewSQLServerSuite.Update_DSI_SQLServer_FinishInstall_ToadDataModeler(ProductName,ProductVersion)
					End Select
					'Update Product Details table data
					Call NewSQLServerSuite.Update_DSI_SQLServer_VerifyProductDetails(ProductName,ProductVersion)
					'Update Verify Reistry table data
					Call NewSQLServerSuite.Update_DSI_SQLServer_VerifyRegistry(ProductName,ProductVersion)
					'Update Silent Install table data
					Call NewSQLServerSuite.Update_SQLServer_SilentInstallMsiBuild(ProductName,ProductVersion)
                                        'Update Shortcut table data
					Call NewSQLServerSuite.Update_DSI_SQLServer_ValidateShortcutAndKeyFile(ProductName,ProductVersion)
                                case UCase("SQLNAVIGATOR")
					Set NewSQLNavigator = New UpdateSQLNavigatorSuite
					'Update all finish installation data
					Select Case Trim(UCase(PreProduct(0)))
						case "SQLNAVIGATOR"
							Call NewSQLNavigator.Update_DSI_FinishInstall_SQLNavigator(ProductName,ProductVersion) 
						case "SQLOPTIMIZERFORORACLE"
							Call NewSQLNavigator.Update_DSI_FinshInstall_OptimizerforOracle(ProductName,ProductVersion)
						case "BENCHMARKFACTORY"
							Call NewSQLNavigator.Update_DSI_FinishInstall_BMF(ProductName,ProductVersion)
                                                case "CODETESTERORACLE"
							Call NewSQLNavigator.Update_DSI_FinishInstall_QuestCodeTester(ProductName,ProductVersion) 
					End Select
					'Update Product Details table data
					   Call NewSQLNavigator.Update_DSI_SQLNavigator_VerifyProductDetail(ProductName,ProductVersion)
					'Update Verify Reistry table data
					   Call NewSQLNavigator.Update_DSI_FinishInstall_VerifyRegistry(ProductName,ProductVersion)
					'Update Silent Install table data
					   Call NewSQLNavigator.Update_SilentInstallMsiBuild(ProductName,ProductVersion)
                                        'Update Shortcut table data
					   Call NewSQLNavigatorSuite.Update_DSI_ValidateShortcutAndKeyFile(ProductName,ProductVersion)
			end Select
		end if
	Next
	
	Conn.Close
	set Conn=Nothing
	
End Sub

Call UpdateTestData()