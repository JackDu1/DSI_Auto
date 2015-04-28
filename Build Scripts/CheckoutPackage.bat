echo branch = %branch%>source.properties
echo TAG_COMMIT_SOURCE = %GIT_COMMIT%>>source.properties
echo CHECKOUT_TIME = %date% %time%>>source.properties
echo ESCROW = %Escrow%>>source.properties

del anchor.txt
del /q *.wip.tmp*
del /q build.properties
del /q display.properties

if not exist w:\ net use /user:FILESERVER\BuildAccess w: \\FILESERVER\Share "FixBoardName!"

if not exist w:\BuildArchives\%branch%.wip set update=true
if not "%GIT_COMMIT%"=="%GIT_PREVIOUS_COMMIT%" set update=true

if not "%update%"=="true" (
   if "%Polling%"=="" "c:\Program Files (x86)\Git\bin\perl" Build\Jenkins\CreateBuildParameters.pl
   if exist build.properties (
     copy build.properties display.properties
   ) else (
     echo branch = %branch% > display.properties
     echo VERSIONCORE = Skipped >> display.properties
   )
   exit 0
)

:wait
if exist w:\BuildArchives\%branch%.lock (
  ping /n 5 127.0.0.1 > nul
  goto :wait
)
echo > w:\BuildArchives\%branch%.packlock
ping /n 1 127.0.0.1 > nul
if exist w:\BuildArchives\%branch%.lock (
  del w:\BuildArchives\%branch%.packlock
  goto :wait
)

w:\Components\7-wip\9.20\7w u -xr!.git -uq0 -twip w:\BuildArchives\%branch%.wip
set failed=%ERRORLEVEL%

del w:\BuildArchives\%branch%.packlock
if NOT %failed%==0 exit %failed%

"c:\Program Files (x86)\Git\bin\perl" Build\Jenkins\CreateBuildParameters.pl

if exist build.properties (
  copy build.properties display.properties
) else (
  echo branch = %branch% > display.properties
  echo VERSIONCORE = Skipped >> display.properties
)