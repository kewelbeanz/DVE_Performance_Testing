' Navigate between modules and tasks inside the modules
' IM to CAIM

' Create WScript object to access filesystem
Set WshShell=WScript.CreateObject("WScript.Shell")

' Start Timer object to count milliseconds
dmlssStartTime=Timer

' Pre-req:  In IM
' Select File menu
WshShell.SendKeys "%F" ' Alt-F
WshShell.SendKeys "{DOWN 6}" ' or however many to go to exit command
WshShell.SendKeys "{ENTER}"  ' Should be a main navigation window

' Open CAIM
WshShell.SendKeys "{DOWN 4}"
WshShell.SendKeys "{ENTER}"




' lengthy contains time as determined by Timer
Dim lengthy2
lengthy2=Round(Timer-dtmStartTime,2)

' Setup for saving the time
Dim output2,fileSystemObject2, filePath2
filePath2="C:\Users\VBASPTHILDER\Documents\PerformanceTesting\ESCS2-1244_Nav_btw_modules\IMToCAIM.txt"
Set fileSystemObject2=CreateObject("Scripting.FileSystemObject")
Set output2=fileSystemObject2.CreateTextFile(filePath2, true)
output2.WriteLine("IM to CAIM took "+lengthy2+" seconds.")
output2.Close
