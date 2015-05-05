C:\Components\Office\setup.exe
ping /n 30 127.0.0.1 > nul

:waitstop
	ping /n 2 127.0.0.1 > nul
	tasklist /v /fi "WINDOWTITLE eq Microsoft Office Professional Plus 2010" |findstr /C:"INFO: No tasks" > nul || goto :waitstop