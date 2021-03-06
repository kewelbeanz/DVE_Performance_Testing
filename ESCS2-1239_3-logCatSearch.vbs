' Log catalog in catalog search in CAIM module

'  Starting from Main Catalog Search screen
' Create WScript Shell Object to access filesystem.
Set WshShell = WScript.CreateObject("WScript.Shell")

' Setup for logging
dtmStartTime=Timer

' Moving to the Scope selection 
WshShell.SendKeys "{Tab}"

'Moving to the log catalog
WshShell.SendKeys "{DOWN}"
WshShell.SendKeys "{DOWN}"
WshShell.SendKeys "{DOWN}"
WshShell.SendKeys "{DOWN}"
WshShell.SendKeys "{DOWN}"

' Moving to Number in Search By field
WshShell.SendKeys "{TAB}"

' Input a valid log catalog number
WshShell.SendKeys "123456"
WshShell.SendKeys "{ENTER}"

' Timer setup
Dim lengthy
lengthy=Round(Timer-dtmStartTime,2)

' Setup for saving the time
Dim output,fileSystemObject, filePath
filePath="C:\Users\VBASPTHILDER\Documents\PerformanceTesting\ESCS2-1238_CAIM_tests\logSearch.txt"
Set fileSystemObject=CreateObject("Scripting.FileSystemObject")
Set output=fileSystemObject.CreateTextFile(filePath, true)
output.WriteLine("Log catalog in CAIM catalog search took "+lengthy+" seconds.")
output.Close
