' Navigate between modules and tasks inside the modules
' System Services to IM

' Create WScript object to access filesystem
Set WshShell=WScript.CreateObject("WScript.Shell")

' Start Timer object to count milliseconds
dmlssStartTime=Timer

' Pre-req:  In System Services
' Select File menu
WshShell.SendKeys "%F" ' Alt-F
WshShell.SendKeys "{DOWN 6}" ' or however many to go to exit command
WshShell.SendKeys "{ENTER}"  ' Should be a main navigation window

' Open IM
WshShell.SendKeys "{DOWN 5}"
WshShell.SendKeys "{ENTER}"




' lengthy contains time as determined by Timer
Dim lengthy2
lengthy2=Round(Timer-dtmStartTime,2)

' Setup for saving the time
Dim output2,fileSystemObject2, filePath2
filePath2="C:\Users\VBASPTHILDER\Documents\PerformanceTesting\ESCS2-1244_Nav_btw_modules\SSToIM.txt"
Set fileSystemObject2=CreateObject("Scripting.FileSystemObject")
Set output2=fileSystemObject2.CreateTextFile(filePath2, true)
output2.WriteLine("System Services to IM took "+lengthy2+" seconds.")
output2.Close
