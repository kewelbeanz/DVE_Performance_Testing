'Sourced Catalog in Catalog Search

' Create WScript Shell Object to access filesystem.
Set WshShell = WScript.CreateObject("WScript.Shell")

' Setup for logging
dtmStartTime=Timer

' Moving to Close button on vertical toolbar
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{ENTER}"

' Moving to the Scope selection 
WshShell.SendKeys "{TAB}"

' Moving to the Sourced Catalog
WshShell.SendKeys "{DOWN}"
WshShell.SendKeys "{ENTER}"

' Moving to Number in Search By field
WshShell.SendKeys "{TAB}"

' Input search term MASKS 
WshShell.SendKeys "{DOWN}"
WshShell.SendKeys "MASKS"
WshShell.SendKeys "{ENTER}"

' Timer setup
Dim lengthy
lengthy=Round(Timer-dtmStartTime,2)

' Setup for saving the time
Dim output,fileSystemObject, filePath
filePath="C:\Users\VBASPTHILDER\Documents\PerformanceTesting\ESCS2-1246_IM_tests2\sourcedCatLogSearch.txt"
Set fileSystemObject=CreateObject("Scripting.FileSystemObject")
Set output=fileSystemObject.CreateTextFile(filePath, true)
output.WriteLine(lengthy)
output.Close
