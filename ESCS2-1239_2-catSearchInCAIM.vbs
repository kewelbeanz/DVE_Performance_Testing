' Cat search in IM Module

' Create WScript Shell Object to access filesystem.
Set WshShell = WScript.CreateObject("WScript.Shell")

' Setup for logging
dtmStartTime=Timer

' The Select a Customer Screen
	' If can cancel out do this
	WshShell.SendKeys "{TAB}"
	WshShell.SendKeys "{TAB}"
	WshShell.SendKeys "{ENTER}"

	' If customer required, do this, replacing 107800 with valid customer id
	WshShell.SendKeys "107800"
	WshShell.SendKeys "{ENTER}"

' Select Navigate
WshShell.SendKeys "%N"

' Select Catalog
WshShell.SendKeys "{DOWN}"
WshShell.SendKeys "{DOWN}"
WshShell.SendKeys "{DOWN}"
WshShell.SendKeys "{ENTER}"

' Input a pre-determined Number in the Search By field
WshShell.SendKeys "123456"
WshShell.SendKeys "{ENTER}"

' Timer setup
Dim lengthy
lengthy=Round(Timer-dtmStartTime,2)

' Setup for saving the time
Dim output,fileSystemObject, filePath
filePath="C:\Users\VBASPTHILDER\Documents\PerformanceTesting\ESCS2-1238_CAIM_tests\catSearch.txt"
Set fileSystemObject=CreateObject("Scripting.FileSystemObject")
Set output=fileSystemObject.CreateTextFile(filePath, true)
output.WriteLine("Catalog Search i IM module took "+lengthy+" seconds.")
output.Close
