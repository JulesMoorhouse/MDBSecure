@Echo off

Echo 1. Clearing Build directory....
Del "MDBSecure\*.*" /Q
Del "MDBSecure\Temp\*.*" /Q
Del "MDBSecure\DistBuild\Temp\*.*" /Q
Del "MDBSecure\Logs\*.*" /Q

Echo 2. Copy Dlls....
copy "DLLs\beside01.exe" "MDBSecure\Beside01.exe" > nul
copy "DLLs\beside02.exe" "MDBSecure\Beside02.exe" > nul    
copy "DLLs\beside03.exe" "MDBSecure\Beside03.exe" > nul    
copy "DLLs\AppBasic.dll" "MDBSecure\AppBasic.dll" > nul
copy "DLLs\WinOnly.dll" "MDBSecure\WinOnly.dll" > nul
copy "DLLs\ProbHand.dll" "MDBSecure\ProbHand.dll" > nul
copy "DLLs\MCLCore.dll" "MDBSecure\MCLCore.dll" > nul
copy "DLLs\UIStyle.dll" "MDBSecure\UIStyle.dll" > nul
copy "..\CodeLibrary\Components\AxInterop.SHDocVw.dll" "MDBSecure\AxInterop.SHDocVw.dll" > nul
copy "..\CodeLibrary\Components\SHDocVw.dll" "MDBSecure\SHDocVw.dll" > nul
copy "DLLs\AppBasic.dll" "MDBSecure\Temp\AppBasic.dll" > nul
copy "DLLs\WinOnly.dll" "MDBSecure\Temp\WinOnly.dll" > nul
copy "DLLs\MCLCore.dll" "MDBSecure\Temp\MCLCore.dll" > nul
copy "DLLs\UIStyle.dll" "MDBSecure\Temp\UIStyle.dll" > nul
copy "..\CodeLibrary\Components\SharpZipLib.dll" "MDBSecure\SharpZipLib.dll" > nul
copy "..\CodeLibrary\Components\AxInterop.SHDocVw.dll" "MDBSecure\Temp\AxInterop.SHDocVw.dll" > nul
copy "..\CodeLibrary\Components\AxInterop.SHDocVw.dll" "MDBSecure\Temp\AxInterop.SHDocVw.dll" > nul

echo .

Echo 3.Compile and copy MDBSecure.exe....
"C:\Program Files (x86)\Microsoft Visual Studio .NET 2003\Common7\IDE\devenv.exe" "..\CodeLibrary\SharewareProjs\MDBSecure\MDBSecure.sln" /rebuild release /out Build\MDBSecure\Logs\MDBSecure.txt
copy "..\CodeLibrary\SharewareProjs\MDBSecure\bin\MDBSecure.exe" "MDBSecure\Temp\MDBSecure.exe"
echo .

echo 4. turn off validation with public keys..
"C:\Program Files (x86)\Microsoft Visual Studio .NET 2003\SDK\v1.1\Bin\sn.exe" -Vr "MDBSecure\Temp\MDBSecure.exe"

echo .

Echo 5. Obfuscate MDBSecure....
echo Fix for MCLCore.dll
copy "..\CodeLibrary\SharewareProjs\MCLCore\bin\MCLCore.dll" "BuildTools\ObfuscatorBook\Executables\MCLCore.dll"
"BuildTools\ObfuscatorBook\Executables\QNDObfuscate.exe" "MDBSecure\Temp\MDBSecure.exe" "MDBSecure\MDBSecure.exe"

echo .

Del "MDBSecure\Temp\*.*" /Q

Echo 6. Signing MDBSecure with MCL Key....
"C:\Program Files (x86)\Microsoft Visual Studio .NET 2003\SDK\v1.1\Bin\sn.exe" -R "MDBSecure\MDBSecure.exe" "..\CodeLibrary\SharewareProjs\IdeasPad\IdeasPad.snk"

echo .

Echo 7. Adding CRC Footer MDBSecure....
"..\CodeLibrary\SharewareProjs\CRCStamp\bin\CRCStamp.exe" "MDBSecure\MDBSecure.exe"
echo .

echo -------- Results --------
type MDBSecure\Logs\MDBSecure.txt
echo -------- Results --------

pause

Echo 8. Compile and copy Setup prog, which will now point at new builds....
"C:\Program Files (x86)\Microsoft Visual Studio .NET 2003\Common7\IDE\devenv.exe" "..\CodeLibrary\SharewareProjs\MDBSecureSetup\MDBSecureSetup.sln" /rebuild release /out MDBSecure\Logs\MDBSecureSetup.txt
copy "..\CodeLibrary\SharewareProjs\MDBSecureSetup\Release\MDBSecure.msi" "MDBSecure\MDBSecure.msi"

echo .
Echo 9. Finished MSI....
echo .

rem Echo Generating backup site(s) update...
rem del D:\Build\BakUpdSites\temp\*.* /Q
rem echo .

rem Echo 10. Compile and copy MCL Keys Dll....
rem "C:\Program Files (x86)\Microsoft Visual Studio .NET 2003\Common7\IDE\devenv.exe" "D:\CodeLibrary\SharewareProjs\MCLKeys\MCLKeys.sln" /rebuild release /out MDBSecure\Logs\MCLKeys.txt
rem copy "D:\CodeLibrary\SharewareProjs\MCLKeys\bin\mclkeys.dll" "MDBSecure\Temp\mclkeys.dll"
rem echo .

rem echo 11. Turn off validation with public keys..
rem "C:\Program Files\Microsoft Visual Studio .NET\FrameworkSDK\Bin\sn.exe" -Vr "MDBSecure\Temp\mclkeys.dll"

rem Echo 12. Obfuscate mclkeys.dll....
rem "D:\Build\BuildTools\ObfuscatorBook\Executables\QNDObfuscate.exe" "MDBSecure\Temp\mclkeys.dll" "MDBSecure\mclkeys.dll"
rem echo .

rem Echo 13. Signing MCL Keys Dll with MCL Key....
rem "C:\Program Files\Microsoft Visual Studio .NET\FrameworkSDK\Bin\sn.exe" -R "MDBSecure\mclkeys.dll" "D:\CodeLibrary\SharewareProjs\IdeasPad\IdeasPad.snk"
rem echo .

rem Echo 14. Zipping and creating MDBSecure Demo updates files....
rem md D:\Build\MCLSoftWeb\temp
rem del D:\Build\MCLSoftWeb\temp\*.* /Q

rem copy "MDBSecure\MDBSecure.exe" "D:\Build\MCLSoftWeb\Temp\MDBSecure.exe"
rem copy "MDBSecure\Beside01.exe" "D:\Build\MCLSoftWeb\Temp\Beside01.exe"
rem copy "MDBSecure\Beside02.exe" "D:\Build\MCLSoftWeb\Temp\Beside02.exe"
rem copy "MDBSecure\mclkeys.dll" "D:\Build\MCLSoftWeb\Temp\mclkeys.dll"
rem echo .

rem D:\Build\BuildTools\ZipUpdCreate.exe D:\Build\MCLSoftWeb\ msd.dat MDBSecure.exe MDBSeDemo.xml http://www.example.com updates/newton/ updates/newton/

rem del D:\Build\MCLSoftWeb\temp\*.* /Q
rem rd D:\Build\MCLSoftWeb\temp
rem echo .

rem Echo Copying backup site base files....
rem copy "MDBSecure\MDBSecure.exe" "D:\Build\BakUpdSites\Temp\MDBSecure.exe"
rem copy "MDBSecure\Beside01.exe" "D:\Build\BakUpdSites\Temp\Beside01.exe"
rem copy "MDBSecure\Beside02.exe" "D:\Build\BakUpdSites\Temp\Beside02.exe"
rem copy "MDBSecure\mclkeys.dll" "D:\Build\BakUpdSites\Temp\mclkeys.dll"
rem echo .

rem echo .
rem echo Finished generating update....
rem echo .

rem Echo Generating Distribution update...
rem echo .
rem copy "D:\CodeLibrary\SharewareProjs\ZipToSelfEx\bin\ZipToSelfEx.exe" "D:\Build\BuildTools\ZipToSelfEx.exe"
copy "MDBSecure\MDBSecure.msi" "MDBSecure\DistBuild\MDBSecure.msi"
rem echo .
rem "D:\Build\BuildTools\ZipToSelfEx.exe" MDBSecure\MDBSecure.exe MDBSe MDBSecure\DistBuild\ D:\Build\MCLSoftWeb\distribution\mdbsecure\
rem echo .
rem echo Copy Distrib to backup dir
rem copy "D:\Build\MCLSoftWeb\distribution\mdbsecure\*.exe" "D:\Build\BakUpdSites\distribution\mdbsecure\*.exe"
rem echo .
rem echo Finished MDBSecure Build, Update and Distrib update
rem echo .

pause

