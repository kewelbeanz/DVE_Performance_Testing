' Catalog search in CAIM

' Create WScript Shell Object to access filesystem.
Set WshShell = WScript.CreateObject("WScript.Shell")

' Setup for logging
dtmStartTime=Timer

'  Select Navigation menu
WshShell.SendKeys "%N"

' Select Catalog Search
WshShell.SendKeys "DOWN 2"
WshShell.SEndKeys "{ENTER}"
 
' Timer setup
Dim lengthy
lengthy=Round(Timer-dtmStartTime,2)

' Setup for saving the time
Dim output,fileSystemObject, filePath
filePath="C:\Users\VBASPTHILDER\Documents\PerformanceTesting\ESCS2-1247_CAIM_tests2\catSearch.txt"
Set fileSystemObject=CreateObject("Scripting.FileSystemObject")
Set output=fileSystemObject.CreateTextFile(filePath, true)
output.WriteLine(lengthy)
output.Close
