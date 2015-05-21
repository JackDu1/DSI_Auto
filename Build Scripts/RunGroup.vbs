 Sub TestDriver()
  On Error Resume Next
   Dim ProjectName
   Dim ProjectSuitPath
   Dim RunGroupName
   Dim TestExecuteApp
   Dim IntegrationObject
   Dim ExcuteTool 
   
   If WScript.Arguments.Count>=4 then
     ExcuteTool=Trim(WScript.Arguments(0))
     ProjectSuitPath=Trim(WScript.Arguments(1))
     ProjectName=Trim(WScript.Arguments(2))
     RunGroupName=Trim(WScript.Arguments(3))
	 
	 ProjectNameArray = Split(ProjectName, ",")
	 RunGroupNameArray = Split(RunGroupName, ",")
   Else
	 wscript.Quit 400
   End If  
 
   if WScript.Arguments(4)<>"" then
	  call UpdateTestParameterTXT()
   End if 

'   Dim OuterCounter, InnerCounter
'   OuterCounter = 0
'   Do
'	   InnerCounter = 0
'	   Set WshShell = WScript.CreateObject("WScript.Shell")
	   'Exec will return an object WshExec
'	   Set oExec = WshShell.Exec("Wscript.exe C:\Automation\License.vbs")
	   
	   set Wshell = WScript.CreateObject("WScript.Shell")
	   if Wshell.ExpandEnvironmentStrings("%programfiles(x86)%")="%programfiles(x86)%" then
			programFiles = "C:/Program Files/"
	   else
			programFiles = "C:/Program Files (x86)/"
	   end if
	   if ExcuteTool="TestComplete" then
		   TestCompletePath = programFiles + "SmartBear/TestComplete 10/Bin/TestComplete.exe"
		   ret = Wshell.run(""""+TestCompletePath+""" /autostart /SilentMode", 0, false)
		  Do
			WScript.Sleep 5000
		    set TestExecuteApp=GetObject(, "TestComplete.TestCompleteApplication.10")
			if Err.Number = 0 or Counter > 5 then
				Exit Do
			end if
			Counter = Counter + 1
			Err.Clear()
		  Loop
		   TestExecuteApp.visible=false
		   Set IntegrationObject = TestExecuteApp.Integration
	   elseif ExcuteTool="TestExecute" then
		  TestExecutePath = programFiles + "SmartBear/TestExecute 10/Bin/TestExecute.exe"
		  ret = Wshell.run(""""+TestExecutePath+""" /autostart /SilentMode", 0, false)
		  Do
			WScript.Sleep 5000
			Set TestExecuteApp=GetObject(, "TestExecute.TestExecuteApplication.10")
			if Err.Number = 0 or Counter > 5 then
				Exit Do
			end if
			Counter = Counter + 1
			Err.Clear()
		  Loop
		  TestExecuteApp.visible=false
		  Set IntegrationObject =TestExecuteApp.Integration
	   else
		  Wscript.Quit 500
	   End if 
	   
	   'Pause 5 second to wait for ExitCode from WshExec
	   WScript.Sleep 5000
	   'To catch ExitCode of WScript.Quit in HandleLicense.vbs
	   If oExec.ExitCode <> 100 Then
		  Wscript.Quit oExec.ExitCode
	   End If

	   IntegrationObject.OpenProjectSuite ProjectSuitPath
	   
	   If Not IntegrationObject.IsProjectSuiteOpened Then
		  TestExecuteApp.Quit
		  Wscript.Quit 1000
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
		If Not IntegrationObject.IsRunning Then
		   IntegrationObject.RunProjectSuite()
		End If
	
	'Wait until TestExecute finished
	While IntegrationObject.IsRunning
      Wscript.sleep(10000)	 
    Wend
	
	Wscript.sleep(10000)	 
    Set TestStatus=IntegrationObject.GetLastResultDescription()
 	TestExecuteApp.Quit
    Wscript.Quit TestStatus.Status
End Sub

Sub UpdateTestParameterTXT()
  
   Dim LocalTxtPath
   Dim Txt
   Dim OpenTxtFile
   Const ForWriting=2
   
   Txt=WScript.Arguments(4)
   ProjectName=Trim(WScript.Arguments(2))
   ProjectNameArray = Split(ProjectName, ",")
   
   for i=0 to Ubound(ProjectNameArray)
	   'Get TestParameters.txt for each project
	   LocalTxtPath="C:\Automation\"+ProjectNameArray(i)+"\TestParameters.txt"
	   Set ObjFso=CreateObject("Scripting.FileSystemobject")
	   'Create file if it does not exist. Otherwise, open it for write only
	   If ObjFso.FileExists(LocalTxtPath) Then
		   Set OpenTxtFile=ObjFso.OpenTextFile(LocalTxtPath,ForWriting,True)
	   Else
		   Set OpenTxtFile=ObjFso.CreateTextFile(LocalTxtPath,true)
	   End If 
	   
	   OpenTxtFile.WriteLine(Txt)
	   OpenTxtFile.Close() 
	   
	   Set  OpenTxtFile=Nothing 
	   Set  ObjFso=Nothing 
	Next
End Sub  

'Kill process by name
Sub KillProcess(strProcessToKill)
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
