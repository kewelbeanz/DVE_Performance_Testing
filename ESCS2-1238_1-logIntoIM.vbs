' Log into IM

' Create WScript object to access filesystem
Set WshShell=WScript.CreateObject("WScript.Shell")

' Start Timer object to count milliseconds
dmlssStartTime=Timer

' Move to IM
WshShell.SendKeys "{DOWN 8}"

' Select IM
WshShell.SendKeys "{ENTER}"

' lengthy contains time as determined by Timer
Dim lengthy
lengthy=Round(Timer-dtmStartTime,2)

' Setup for saving the time
Dim output,fileSystemObject, filePath
filePath="C:\Users\VBASPTHILDER\Documents\PerformanceTesting\ESCS2-1238_IM_tests\logIntoCaim.txt"
Set fileSystemObject=CreateObject("Scripting.FileSystemObject")
Set output=fileSystemObject.CreateTextFile(filePath, true)
output.WriteLine(lengthy)
output.Close