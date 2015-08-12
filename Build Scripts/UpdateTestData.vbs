'Option Explicit

Dim Conn

Class UpdateOracleSuite

	'==============================DSI_FinishInstall_ToadforOracle========================================
	Function Update_DSI_FinishInstall_ToadforOracle(ByVal StrProduct,ByVal StrVersion)
		
		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
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
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_ToadforOracle=True
		else
			Update_DSI_FinishInstall_ToadforOracle=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinshInstall_OptimizerforOracle========================================
	Function Update_DSI_FinshInstall_OptimizerforOracle(ByVal StrProduct,ByVal StrVersion)
		
		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
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
		
		if Err.Number = 0 then
			Update_DSI_FinshInstall_OptimizerforOracle=True
		else
			Update_DSI_FinshInstall_OptimizerforOracle=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_BMF========================================
	Function Update_DSI_FinishInstall_BMF(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
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
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_BMF set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Benchmark Factory%" + StrProduct +"'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_FinishInstall_BMF where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Benchmark Factory%" + StrProduct +"'"
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
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_BMF set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Benchmark Factory%" + StrProduct +"'"
		
		Rec.Close
		Set Rec	= Nothing
		if Err.Number = 0 then
			Update_DSI_FinishInstall_BMF=True
		else
			Update_DSI_FinishInstall_BMF=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_SpotlightonOracle========================================
	Function Update_DSI_FinishInstall_SpotlightonOracle(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
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
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_SpotlightonOracle=True
		else
			Update_DSI_FinishInstall_SpotlightonOracle=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_ToadDataModeler========================================
	Function Update_DSI_FinishInstall_ToadDataModeler(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_ToadDataModeler=false
			wscript.quit 100
		end if
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_ToadDataModeler set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Toad_ Data Modeler'"
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_FinishInstall_ToadDataModeler where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Toad_ Data Modeler'"
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
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_ToadDataModeler set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Toad_ Data Modeler'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_ToadDataModeler=True
		else
			Update_DSI_FinishInstall_ToadDataModeler=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_QuestCodeTester========================================
	Function Update_DSI_FinishInstall_QuestCodeTester(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_QuestCodeTester=false
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
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_QuestCodeTester=True
		else
			Update_DSI_FinishInstall_QuestCodeTester=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_BackupReportForOracle========================================
	Function Update_DSI_FinishInstall_BackupReportForOracle(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_BackupReportForOracle=false
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
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_BackupReportForOracle=True
		else
			Update_DSI_FinishInstall_BackupReportForOracle=False
			Err.Clear
		end if

	End Function


	'==============================DSI_FinishInstall_BackupReportForOracle========================================
	Function Update_DSI_FinishInstall_ToadforMySQLFreeware(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_ToadforMySQLFreeware=false
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
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_ToadforMySQLFreeware=True
		else
			Update_DSI_FinishInstall_ToadforMySQLFreeware=False
			Err.Clear
		end if

	End Function
	
	'==============================DSI_ProductSelectionPage_VerifyProductDetail========================================
	Function Update_DSI_ProductSelectionPage_VerifyProductDetail(ByVal StrProduct,ByVal StrVersion)
		
		on error resume next
		
		if IsEmpty(StrProduct) then
			Update_DSI_ProductSelectionPage_VerifyProductDetail=false
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
					StrProduct="TOAD% DATA MODELER"
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
		
		if Err.Number = 0 then
			Update_DSI_ProductSelectionPage_VerifyProductDetail=True
		else
			Update_DSI_ProductSelectionPage_VerifyProductDetail=False
			Err.Clear
		end if

	End Function
	
	'==============================DSI_DSI_FinishInstall_VerifyRegistry========================================
	Function Update_DSI_FinishInstall_VerifyRegistry(ByVal StrProduct,ByVal StrVersion)
		
		on error resume next
		
		if IsEmpty(StrProduct) then
			Update_DSI_ProductSelectionPage_VerifyProductDetail=false
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
					StrProduct="TOAD% DATA MODELER"
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
		
		Conn.Execute "Update DSI.dbo.DSI_Oracle_VerifyRegistry set  I_ProductVersion =" + "'" + StrVersion + "'" + " where Projectid = 1 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_InstallerDisplayProductName) like '" + StrProduct + "'"
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_VerifyRegistry=True
		else
			Update_DSI_FinishInstall_VerifyRegistry=False
			Err.Clear
		end if

	End Function

End Class

Class UpdateSAPSuite

	'==============================DSI_FinishInstall_ToadforSybase========================================
	Function Update_DSI_FinishInstall_ToadforSybase(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_ToadforSybase=false
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
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_ToadforSybase=True
		else
			Update_DSI_FinishInstall_ToadforSybase=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_QuestSQLOptimizerforSybase========================================
	Function Update_DSI_FinishInstall_QuestSQLOptimizerforSybase(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinshInstall_OptimizerforOracle=false
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
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_QuestSQLOptimizerforSybase=True
		else
			Update_DSI_FinishInstall_QuestSQLOptimizerforSybase=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_BMF========================================
	Function Update_DSI_FinishInstall_BMF(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
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
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_SAP_BMF set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Benchmark Factory%" + StrProduct + "'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_SAP_BMF where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Benchmark Factory%'"
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
		
		Conn.Execute "Update DSI.dbo.DSI_SAP_BMF set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Benchmark Factory%'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_BMF=True
		else
			Update_DSI_FinishInstall_BMF=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_SpotlightonSybase========================================
	Function Update_DSI_FinishInstall_SpotlightonSybase(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_SpotlightonSybase=false
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
		StrVer 		= 	StrMainVer(0) + "." + StrMainVer(1) + "." + StrMainVer(2)
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		StrColFolder 	= 	regEx.Replace(StrColFolder,StrVer)
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_SpotlightonSybase set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Spotlight% on SAP% ASE'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_SpotlightonSybase=True
		else
			Update_DSI_FinishInstall_SpotlightonSybase=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_ToadDataModeler========================================
	Function Update_DSI_FinishInstall_ToadDataModeler(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_ToadDataModeler=false
			wscript.quit 100
		end if
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_SAP_ToadDataModeler set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Toad% Data Modeler'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_SAP_ToadDataModeler where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Toad% Data Modeler'"
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
		
		Conn.Execute "Update DSI.dbo.DSI_SAP_ToadDataModeler set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Toad% Data Modeler'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_ToadDataModeler=True
		else
			Update_DSI_FinishInstall_ToadDataModeler=False
			Err.Clear
		end if

	End Function
	
	'==============================DSI_ProductSelectionPage_VerifyProductDetails========================================
	Function Update_DSI_ProductSelectionPage_VerifyProductDetails(ByVal StrProduct,ByVal StrVersion)
				
		on error resume next
		
		if IsEmpty(StrProduct) then
			Update_DSI_ProductSelectionPage_VerifyProductDetails=false
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "TOADFORSAP_X86_EN"
					StrProduct="TOAD_ FOR SAP SOLUTIONS"
				case "TOADDATAMODELER_X86_EN"
					StrProduct="TOAD_ DATA MODELER"
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
		
		Conn.Execute "Update DSI.dbo.DSI_SAP_VerifyProductDetails set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_ProductName) like '" + StrProduct + "'"
		
		if Err.Number = 0 then
			Update_DSI_ProductSelectionPage_VerifyProductDetails=True
		else
			Update_DSI_ProductSelectionPage_VerifyProductDetails=False
			Err.Clear
		end if

	End Function
	
	'==============================DSI_FinishInstall_VerifyRegistry========================================
	Function Update_DSI_FinishInstall_VerifyRegistry(ByVal StrProduct,ByVal StrVersion)
		
		on error resume next
		
		if IsEmpty(StrProduct) then
			Update_DSI_ProductSelectionPage_VerifyProductDetails=false
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "TOADFORSAP_X86_EN"
					StrProduct="TOAD_ FOR SAP SOLUTIONS"
				case "TOADDATAMODELER_X86_EN"
					StrProduct="TOAD_ DATA MODELER"
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
		
		Conn.Execute "Update DSI.dbo.DSI_SAP_VerifyRegistry set  I_ProductVersion =" + "'" + StrVersion + "'" + " where Projectid = 3 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_InstallerDisplayProductName) like '" + StrProduct + "'"
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_VerifyRegistry=True
		else
			Update_DSI_FinishInstall_VerifyRegistry=False
			Err.Clear
		end if

	End Function

End Class

Class UpdateDB2Suite

	'==============================DSI_FinishInstall_ToadforIBMDB2LUW========================================
	Function Update_DSI_FinishInstall_ToadforIBMDB2LUW(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_ToadforIBMDB2LUW=false
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "TOADFORDB2_X86_EN"
					StrProduct="Toad% for IBM% DB2%"
				case "TOADFORDB2_TRIAL_X86_EN"
					StrProduct="Toad% for IBM% DB2% Trial"
			end select	
		end if
		wscript.echo(StrProduct)
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
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_ToadforIBMDB2LUW=True
		else
			Update_DSI_FinishInstall_ToadforIBMDB2LUW=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_QuestSQLOptimizerforIBMDB2========================================
	Function Update_DSI_FinishInstall_QuestSQLOptimizerforIBMDB2(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_QuestSQLOptimizerforIBMDB2=false
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
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_QuestSQLOptimizerforIBMDB2=True
		else
			Update_DSI_FinishInstall_QuestSQLOptimizerforIBMDB2=False
			Err.Clear
		end if

	End Function
	
	'==============================DSI_FinishInstall_QuestSQLOptimizerForDB2zOS========================================
	Function Update_DSI_FinishInstall_QuestSQLOptimizerForDB2zOS(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_QuestSQLOptimizerForDB2zOS=false
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
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_QuestSQLOptimizerForDB2zOS=True
		else
			Update_DSI_FinishInstall_QuestSQLOptimizerForDB2zOS=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_BMF========================================
	Function Update_DSI_FinishInstall_BMF(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
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
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_DB2_BMF set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Benchmark Factory%" + StrProduct + "'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_DB2_BMF where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Benchmark Factory%" + StrProduct + "'"
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
		
		Conn.Execute "Update DSI.dbo.DSI_DB2_BMF set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Benchmark Factory%" + StrProduct + "'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_BMF=True
		else
			Update_DSI_FinishInstall_BMF=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_SpotlightonIBMDB2========================================
	Function Update_DSI_FinishInstall_SpotlightonIBMDB2(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_SpotlightonIBMDB2=false
			wscript.quit 100	
		end if
		
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_SpotlightonIBMDB2 set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Spotlight_ on IBM_ DB2_ LUW'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_FinishInstall_SpotlightonIBMDB2 where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Spotlight_ on IBM_ DB2_ LUW'"
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
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_SpotlightonIBMDB2 set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Spotlight_ on IBM_ DB2_ LUW'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_SpotlightonIBMDB2=True
		else
			Update_DSI_FinishInstall_SpotlightonIBMDB2=False
			Err.Clear
		end if

	End Function

	'==============================DSI_FinishInstall_ToadDataModeler========================================
	Function Update_DSI_FinishInstall_ToadDataModeler(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_ToadDataModeler=false
			wscript.quit 100	
		end if
		
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_DB2_ToadDataModeler set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Toad% Data Modeler'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_DB2_ToadDataModeler where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Toad% Data Modeler'"
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
		
		Conn.Execute "Update DSI.dbo.DSI_DB2_ToadDataModeler set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Toad% Data Modeler'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_ToadDataModeler=True
		else
			Update_DSI_FinishInstall_ToadDataModeler=False
			Err.Clear
		end if

	End Function
	
	'==============================DSI_ProductSelectionPage_VerifyProductDetails========================================
	Function Update_DSI_ProductSelectionPage_VerifyProductDetails(ByVal StrProduct,ByVal StrVersion)
		
		on error resume next
		
		if IsEmpty(StrProduct) then
			Update_DSI_ProductSelectionPage_VerifyProductDetails=false
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "TOADFORDB2_X86_EN"
					StrProduct="TOAD_ FOR IBM_ DB2_"
				case "TOADFORDB2_TRIAL_X86_EN"
					StrProduct="TOAD_ FOR IBM_ DB2_ TRIAL"
				case "TOADDATAMODELER_X86_EN"
					StrProduct="TOAD_ DATA MODELER"
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
		
		if Err.Number = 0 then
			Update_DSI_ProductSelectionPage_VerifyProductDetails=True
		else
			Update_DSI_ProductSelectionPage_VerifyProductDetails=False
			Err.Clear
		end if

	End Function
	
	'==============================DSI_FinishInstall_VerifyRegistry========================================
	Function Update_DSI_FinishInstall_VerifyRegistry(ByVal StrProduct,ByVal StrVersion)
		
		on error resume next
		
		if IsEmpty(StrProduct) then
			Update_DSI_FinishInstall_VerifyRegistry=false
			wscript.quit 100
		else
			select case UCase(StrProduct)
				case "TOADFORDB2_X86_EN"
					StrProduct="TOAD_ FOR IBM_ DB2_"
				case "TOADFORDB2_TRIAL_X86_EN"
					StrProduct="TOAD_ FOR IBM_ DB2_ TRIAL"
				case "TOADDATAMODELER_X86_EN"
					StrProduct="TOAD_ DATA MODELER"
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
		
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_VerifyRegistry set  I_ProductVersion =" + "'" + StrVersion + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and UPPER(I_InstallerDisplayProductName) like '" + StrProduct + "'"
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_VerifyRegistry=True
		else
			Update_DSI_FinishInstall_VerifyRegistry=False
			Err.Clear
		end if

	End Function

End Class

Class UpdateSQLServerSuite

	'==============================DSI_DSI_FinishInstall_BMF========================================
	Function Update_DSI_FinishInstall_BMF(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
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
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_BMF set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Benchmark Factory%" + StrProduct +"'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_DB2_ToadDataModeler where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Toad% Data Modeler'"
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
		
		Conn.Execute "Update DSI.dbo.DSI_DB2_ToadDataModeler set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Toad% Data Modeler'"
		
		Rec.Close
		Set Rec	= Nothing
		
		if Err.Number = 0 then
			Update_DSI_FinishInstall_BMF=True
		else
			Update_DSI_FinishInstall_BMF=False
			Err.Clear
		end if

	End Function


	'==============================DSI_FinishInstall_ToadDataModeler========================================
	Function Update_DSI_FinishInstall_ToadDataModeler(ByVal StrProduct,ByVal StrVersion)

		Dim StrColFolder,StrMainVer,Query,StrVer
		on error resume next
		
		Set regEx = New RegExp
		
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
		'Update I_Version Column
		Conn.Execute "Update DSI.dbo.DSI_FinishInstall_ToadDataModeler set  I_Version =" + "'" + StrVersion + "'" + " where Projectid = 4 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Toad% Data Modeler'"
		
		'Update I_InstallFolder Column Record
		Set Rec		=	CreateObject("ADODB.Recordset")
		Query		= 	"Select I_InstallFolder from DSI.dbo.DSI_DB2_ToadDataModeler where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Toad% Data Modeler'"
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
		
		Conn.Execute "Update DSI.dbo.DSI_DB2_ToadDataModeler set  I_InstallFolder =" + "'" + StrColFolder + "'" + " where Projectid = 2 and UPPER(I_AutoUpdate) = 'TRUE' and I_ProductName like 'Toad% Data Modeler'"
		
		Rec.Close
		Set Rec	= Nothing
		
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

	Dim XMLDoc,FSO,regEx
	Dim ErrorMsg,i
	Dim NewOracleSuite,NewSAPSuite,NewDB2Suite,NewSQLServerSuite
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
	
	'wscript.echo("there are total: " + cstr(RootNode.childnodes.length) + " Products in this build")
	
	Set Conn	=	CreateObject("ADODB.Connection")
	Conn.Open "Driver=SQL Server;Server=10.6.208.62;Database=DSI;uid=sa;pwd=Quest6848;"	
	
	set ProductNode	=	Productnode.item(0)
	set regEx		= 	New RegExp
	
	For i = 0 to RootNode.childnodes.length - 1
		
		NodeName	=	productnode.childnodes.item(i).text
		NodeName	=	Split(NodeName,"=")
			
		regEx.Pattern 	= 	"\d+(\.\d+)+"
		regEx.Global	=	True
		set Matches 	=	regEx.Execute(NodeName(1))
			
		for each match in Matches
			ProductVersion	=	match.value
		next
		ProductName	=	Mid(NodeName(0),1,Len(NodeName(0)) - 15)
		'ProductName=Split(NodeName(0),"PACKAGE")
		
		if ProductName <> "" and ProductVersion <> "" then
			PreProduct=Split(ProductName,"_")
			Select Case UCase(StrProduct)
				case UCase("ORACLE")
					Set NewOracleSuite	=	New UpdateOracleSuite
					Select Case Trim(UCase(PreProduct(0)))
						case "TOADFORORACLE"
							if NewOracleSuite.Update_DSI_FinishInstall_ToadforOracle(ProductName,ProductVersion) then
								'wscript.echo("Update DSI_FinishInstall_ToadforOracle table successful!")
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
					if NewOracleSuite.Update_DSI_ProductSelectionPage_VerifyProductDetail(ProductName,ProductVersion) then
						'wscript.echo("Update DSI_ProductSelectionPage_VerifyProductDetail table successful!")
					end if
					if NewOracleSuite.Update_DSI_FinishInstall_VerifyRegistry(ProductName,ProductVersion) then
						'wscript.echo("Update DSI_FinishInstall_VerifyRegistry table successful!")
					end if
				case UCase("DB2")
					Set NewDB2Suite	=	New UpdateDB2Suite
					Select Case Trim(UCase(PreProduct(0)))
						case "TOADFORDB2"
							Call NewDB2Suite.Update_DSI_FinishInstall_ToadforIBMDB2LUW(ProductName,ProductVersion)
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
					if NewDB2Suite.Update_DSI_ProductSelectionPage_VerifyProductDetails(ProductName,ProductVersion) then
						'wscript.echo("Update DSI_ProductSelectionPage_VerifyProductDetail table successful!")
					end if
					if NewDB2Suite.Update_DSI_FinishInstall_VerifyRegistry(ProductName,ProductVersion) then
						'wscript.echo("Update DSI_FinishInstall_VerifyRegistry table successful!")
					end if
				case UCase("SAP")
					Set NewSAPSuite	=	New UpdateSAPSuite
					Select Case Trim(UCase(PreProduct(0)))
						case "TOADFORSAP"
							if NewSAPSuite.Update_DSI_FinishInstall_ToadforSybase(ProductName,ProductVersion) then
								'wscript.echo("Update DSI_FinishInstall_ToadforSybase table successful!")
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
					if NewSAPSuite.Update_DSI_ProductSelectionPage_VerifyProductDetails(ProductName,ProductVersion) then
						'wscript.echo("Update DSI_ProductSelectionPage_VerifyProductDetail table successful!")
					end if
					if NewSAPSuite.Update_DSI_FinishInstall_VerifyRegistry(ProductName,ProductVersion) then
						'wscript.echo("Update DSI_FinishInstall_VerifyRegistry table successful!")
					end if
				case UCase("SQLSERVER")
					Set NewSQLServerSuite	=	New UpdateSQLSERVERSuite
					'not implemented
			end Select
		end if
	Next
	
	Conn.Close
	set Conn=Nothing
	
End Sub

Call UpdateTestData()