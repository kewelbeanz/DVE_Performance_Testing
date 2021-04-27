' Select UP Assign (User Privileges) in System Services

' From Navigation window, go to System Services
' Create WScript object to access filesystem
Set WshShell=WScript.CreateObject("WScript.Shell")

' Start Timer object to count milliseconds
dmlssStartTime=Timer

' Move to System Services
WshShell.SendKeys "{DOWN 8}"
' WshShell.SendKeys "{DOWN}"
' WshShell.SendKeys "{DOWN}"
' WshShell.SendKeys "{DOWN}"
' WshShell.SendKeys "{DOWN}"
' WshShell.SendKeys "{DOWN}"
' WshShell.SendKeys "{DOWN}"
' WshShell.SendKeys "{DOWN}"

WshShell.SendKeys "{ENTER}"

' lengthy contains time as determined by Timer
Dim lengthy
lengthy=Round(Timer-dtmStartTime,2)

' Setup for saving the time
Dim output,fileSystemObject, filePath
filePath="C:\Users\VBASPTHILDER\Documents\PerformanceTesting\ESCS2-1240_System_Services\logIntoSysServ.txt"
Set fileSystemObject=CreateObject("Scripting.FileSystemObject")
Set output=fileSystemObject.CreateTextFile(filePath, true)
output.WriteLine(lengthy)
output.Close

' Navigate menu - maybe tab tab 
WshShell.SendKeys "{%N}"

' Go to User Priv Assign
WshShell.SendKeys "{DOWN 16}"
' WshShell.SendKeys "{DOWN}"
' WshShell.SendKeys "{DOWN}"
' WshShell.SendKeys "{DOWN}"
' WshShell.SendKeys "{DOWN}"
' WshShell.SendKeys "{DOWN}"
' WshShell.SendKeys "{DOWN}"
' WshShell.SendKeys "{DOWN}"
' WshShell.SendKeys "{DOWN}"
' WshShell.SendKeys "{DOWN}"
' WshShell.SendKeys "{DOWN}"
' WshShell.SendKeys "{DOWN}"
' WshShell.SendKeys "{DOWN}"
' WshShell.SendKeys "{DOWN}"
' WshShell.SendKeys "{DOWN}"
' WshShell.SendKeys "{DOWN}"
' WshShell.SendKeys "{ENTER}"

' lengthy contains time as determined by Timer
Dim lengthy2
lengthy2=Round(Timer-dtmStartTime,2)

' Setup for saving the time
Dim output2,fileSystemObject2, filePath2
filePath2="C:\Users\VBASPTHILDER\Documents\PerformanceTesting\ESCS2-1240_System_Services\selectUpAssign.txt"
Set fileSystemObject2=CreateObject("Scripting.FileSystemObject")
Set output2=fileSystemObject2.CreateTextFile(filePath2, true)
output2.WriteLine("Select User Privileges in System Service took "+lengthy2+" seconds.")
output2.Close
