' Create WScript Shell Object to access file system
Set WshShell=Wscript.CreateObject("WScript.Shell")

' Start Timer object to count milliseconds
dmlssStartTime=Timer

' Start / Run DMLSS - might need to adjust if actual app name is different
WshShell.Run "%windir%\dmlss.exe" 

' Select, or bring focus to a window named `DMLSS`
WshShell.AppActivate "DMLSS"

' Wait for 5 seconds to account for processor speed
WScript.Sleep 5000

' How long did it take?
' Results stored in a variable named `lengthy`
' Dim sets aside memory space for variable
' Round rounds the result to the specified number of characters
Dim lengthy
lengthy=Round(Timer-dmlssStartTime,2)

' Removing the 5 seconds added earlier for processor speed
lengthy=Round(lengthy-5,2)

' Setup for saving the time
' First, reserve space in memory for file objects
Dim output, fileSystemObject, filePath

' Set the filepath to the user's location
filePath="C:\Users\Public\Public Documents\ESCS2_1236_test.txt"

' Create the fileSystemObject to create the output TextFile object
Set fileSystemObject=CreateObject("Scripting.FileSystemObject")
Set output=fileSystemObject.CreateTextFile(filePath, true)

' Write the time contained in the `lengthy` object to the TextFile found at the filePath destination
output.WriteLine(lengthy)

' Close the filesystem
output.Close
