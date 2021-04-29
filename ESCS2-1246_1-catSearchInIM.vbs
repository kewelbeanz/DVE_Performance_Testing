' Catalog Search in IM Module

' Pre-req - in IM

' Create WScript object to access filesystem
Set WshShell=WScript.CreateObject("WScript.Shell")

' Start Timer object to count milliseconds
dmlssStartTime=Timer

' Get Navigate menu
WshShell.SendKeys "%N"

' Go to Catalog Search
WshShell.SendKeys "{DOWN 3}"

' Select Catalog Search
WshShell.SendKeys "{ENTER}"



' lengthy contains time as determined by Timer
Dim lengthy2
lengthy2=Round(Timer-dtmStartTime,2)

' Setup for saving the time
Dim output2,fileSystemObject2, filePath2
filePath2="C:\Users\VBASPTHILDER\Documents\PerformanceTesting\ESCS2-1246_IM_Tests2\catSearchInIM.txt"
Set fileSystemObject2=CreateObject("Scripting.FileSystemObject")
Set output2=fileSystemObject2.CreateTextFile(filePath2, true)
output2.WriteLine("Catalog Search in IM Module took "+lengthy2+" seconds.")
output2.Close
