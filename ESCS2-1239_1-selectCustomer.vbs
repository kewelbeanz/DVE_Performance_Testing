' Select Customer in CAIM Module

' Create WScript Shell Object to access filesystem.
Set WshShell = WScript.CreateObject("WScript.Shell")

' Setup for logging
dtmStartTime=Timer

' The Select a Customer Screen
' replace 107800 with valid customer id
	WshShell.SendKeys "107800"
	WshShell.SendKeys "{ENTER}"
' Timer setup
Dim lengthy
lengthy=Round(Timer-dtmStartTime,2)

' Setup for saving the time
Dim output,fileSystemObject, filePath
filePath="C:\Users\VBASPTHILDER\Documents\PerformanceTesting\ESCS2-1239_CAIM_tests\catSearch.txt"
Set fileSystemObject=CreateObject("Scripting.FileSystemObject")
Set output=fileSystemObject.CreateTextFile(filePath, true)
output.WriteLine("Select Customer in CAIM module took "+lengthy+" seconds.")
output.Close
