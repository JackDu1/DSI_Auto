 Sub TestDriver()
  On Error Resume Next
   Dim ProjectName
   Dim ProjectSuitPath
   Dim RunGroupName
   Dim TestExecuteApp
   Dim IntegrationObject
   Dim ExecuteTool 
   
   If WScript.Arguments.Count>=4 then
     ExecuteTool=Trim(WScript.Arguments(0))
     ProjectSuitPath=Trim(WScript.Arguments(1))
     ProjectName=Trim(WScript.Arguments(2))
     RunGroupName=Trim(WScript.Arguments(3))
	 
	 ProjectNameArray = Split(ProjectName, ",")
	 RunGroupNameArray = Split(RunGroupName, ",")
   Else
	 wscript.Quit 400
   End If
	   
	   set Wshell = WScript.CreateObject("WScript.Shell")
	   if Wshell.ExpandEnvironmentStrings("%programfiles(x86)%")="%programfiles(x86)%" then
			programFiles = "C:/Program Files/"
	   else
			programFiles = "C:/Program Files (x86)/"
	   end if
	   if ExecuteTool="TestComplete" or ExecuteTool="TestExecute" then
		   ExecuteToolPath = programFiles + "SmartBear/" + ExecuteTool + " 10/Bin/" + ExecuteTool + ".exe"
		   ret = Wshell.run(""""+ExecuteToolPath+""" /autostart /SilentMode", 0, false)
		   'Wait TestExecute launch successful (2min)
		   For i=1 to 120
				if CheckApplicationIsRun(ExecuteTool+".exe") then
					WScript.Echo("Launching "+ExecuteTool+" succeeded!")
					Exit For
				End if
				if i>=120 then
					WScript.Echo("Launching "+ExecuteTool+" failed!")
					Wscript.Quit 300
				End if
				WScript.Sleep 1000
		   Next
		   'Wait to Get TestExecute Object successful (20mins)
		   'Script cannot get the object if there is no license available
		  Do
			WScript.Sleep 10000
		    set TestExecuteApp=GetObject(, ExecuteTool+"."+ExecuteTool+"Application")
			if Err.Number = 0 or Counter > 120 then
				Exit Do
			end if
			Counter = Counter + 1
			Err.Clear()
		  Loop
		  if Err.Number <> 0 or VarType(TestExecuteApp)<>9 then
			'VarType=9 means getobject successfully, else failed.
			WScript.Echo("Object type : " + cstr(VarType(TestExecuteApp)))
			WScript.Echo("Error number : " + cstr(Err.Number))
			WScript.Echo("Cannot get TestExecute object.")
			KillProcess(ExecuteTool+".exe")
			Wscript.Quit 400
		  end if
		   WScript.Echo("GetObject successfully.")
		   TestExecuteApp.visible=false
		   Set IntegrationObject = TestExecuteApp.Integration
	   else
		  Wscript.Quit 500
	   End if 
	   
	   'Wait to Get TestExecute Object successful (2mins)
	   'Open project suite
	   IntegrationObject.OpenProjectSuite ProjectSuitPath
	   for i=0 to 12
		   If Not IntegrationObject.IsProjectSuiteOpened Then
			  WScript.Sleep 10000
			  IntegrationObject.OpenProjectSuite ProjectSuitPath
		   End If
	   Next
	   If Not IntegrationObject.IsProjectSuiteOpened Then
		  TestExecuteApp.Quit
		  Wscript.Quit 1000
		 else
			WScript.Echo("Opened project suite successfully.")
	   End If

		'Select the projects in project suite
		Set ProjectSuiteItems=IntegrationObject.TestSuite("")
		ProjectCount = ProjectSuiteItems.Count
		for i=0 to ProjectCount-1
			Set ProjectItem = ProjectSuiteItems.TestItem(i)
			for each x in ProjectNameArray	
				if strcomp(ProjectItem.ProjectName, x)=0 then
					ProjectItem.Enabled=True
					Exit for
				else
					ProjectItem.Enabled=False
				End if
			Next
		Next
		
		'Select test items in each project
		for i=0 to Ubound(ProjectNameArray)
			Set ProjectItems=IntegrationObject.TestSuite(ProjectNameArray(i))
			ItemsCount=ProjectItems.Count
			RunGroupNameArrayOfEachProject = Split(RunGroupNameArray(i), "&")
			for j=0 to ItemsCount-1
				set TestItem=ProjectItems.TestItem(j)
				for each x in RunGroupNameArrayOfEachProject
					If strcomp(TestItem.Name,x)=0 then
						TestItem.Enabled=True  
						Exit for
					else
						TestItem.Enabled=False
					End If
				Next
			Next
		Next
		
		'Start Sentinel HASP License Manager service if it is stopped
		StartService("hasplms")
		
		'Start to run		
		'If Not IntegrationObject.IsRunning Then
		'   IntegrationObject.RunProjectSuite()
		'End If
	
	'Wait until TestExecute finished
	'While IntegrationObject.IsRunning
     'Wscript.sleep(10000)
   'Wend
	
	'Wscript.sleep(10000)	 
   'Set TestStatus=IntegrationObject.GetLastResultDescription()
	TestExecuteApp.Quit
   Wscript.Quit TestStatus.Status
End Sub

'Kill process by name
Sub KillProcess(strProcessToKill)
	WScript.Echo("Kill Process: "+ strProcessToKill)
	Dim WMI, ProcessList
	Set WMI = GetObject("WinMgmts:\\.\root\cimv2")
	Set ProcessList = WMI.ExecQuery("select * from win32_process where name='" & strProcessToKill & "'")
	For Each process In ProcessList
		process.Terminate()
	Next
	Set ProcessList = Nothing
	Set WMI = Nothing
End Sub

'Start service by name
Sub StartService(svcToStart)
	WScript.Echo("Start Service : " + svcToStart)
	Dim WMI, svcs
	Set WMI = GetObject("WinMgmts:\\.\root\cimv2")
	Set svcs = WMI.ExecQuery("select * from win32_service where name = '" & svcToStart & "'")
	For Each svc In svcs
		'Start service only when it is stopped
		If svc.state = "Stopped" Then
			svc.StartService()
			WScript.Sleep 5000
		End If
	Next
End Sub

'To check if an application is running 
Function CheckApplicationIsRun(ByVal processName)
	Dim WMI, Processes, Exist
	CheckApplicationIsRun = False
	Set WMI = GetObject("WinMgmts:\\.")
	Set Processes = WMI.InstancesOf("Win32_Process")
	For Each process In Processes
		If process.Name = processName Then
			CheckApplicationIsRun = True
			Set Processes = Nothing
			Set WMI = Nothing
			Exit Function
		End If
	Next
End Function 

Call TestDriver()
