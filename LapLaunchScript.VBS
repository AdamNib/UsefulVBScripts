'-------------------------------------------------------------------------------------------------------------------------------------
'Laptop Launch Script
'AJN Spet 2014
'What it does - Checks for a connection to the network (ie is Wifi working?), Prompts for userlogin (as we always use the scilaptop user account) writes out login details to lgo file, Logs into AD and sets up an N drive - details progress...
'-------------------------------------------------------------------------------------------------------------------------------------
'Change Log
'Detail any changes to the script below in the format - Who? | When? | What?
'
'
'
'
'
'
'-------------------------------------------------------------------------------------------------------------------------------------

'Main Programme starts here.

Option Explicit
'on error resume next

'Constants and Variable setup


Dim strLoginName,strPassword, strServer, strVLogDir,strVLogFile, strComputerName, StrLoginFailure,objFSO,ie,objNetwork,objTrans,strDomain

'Objects used throughout
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set ie = WScript.CreateObject("InternetExplorer.Application","IE_") 
Set objNetwork = CreateObject("WScript.Network")
Set objTrans   =   CreateObject("NameTranslate")

'Variables used Throughout

strServer = "uppingham.co.uk" 'OK are we connected to the network
strVLogDir = "C:\TESTING123\"
StrLoginFailure = ""
strDomain = objNetwork.UserDomain
msgbox strDomain
'-----------------------------------------------------------------------------------
'Main Program starts here'
'-----------------------------------------------------------------------------------


InitLogFile			'Get the log file ready
Main

Sub Main

if Willitping(strServer) then
	WriteLog("Attached to the network Fine")
	if Loginbox then
		MapHomeDrive
	else
		StrLoginFailure="Sorry, Username / Password incorrect..."
		Main
	end if
	
else
	WriteLog("WiFi not working on login - could not communicate with DC")
	msgbox "Sorry, couldn't connect to the School Wifi working in Off-line mode."
end if
	
end sub


'-------------------------------------------------------------------------------------------------------------------------------------
' WillitPing - This function returns True if the specified host could be pinged.
' myHostName can be a computer name or IP address.
' The Win32_PingStatus class used in this function requires Windows XP or later.
' This function is based on the TestPing function in a sample script by Don Jones
' http://www.scriptinganswers.com/vault/computer%20management/default.asp#activedirectoryquickworkstationinventorytxt
'-------------------------------------------------------------------------------------------------------------------------------------
Function Willitping( myHostName )

    ' Standard housekeeping
    Dim colPingResults, objPingResult, strQuery, Ping

    ' Define the WMI query
    strQuery = "SELECT * FROM Win32_PingStatus WHERE Address = '" & myHostName & "'"

    ' Run the WMI query
    Set colPingResults = GetObject("winmgmts://./root/cimv2").ExecQuery( strQuery )

    ' Translate the query results to either True or False
    For Each objPingResult In colPingResults
        If Not IsObject( objPingResult ) Then
            Willitping = False
        ElseIf objPingResult.StatusCode = 0 Then
            Willitping = True
        Else
            Willitping = False
        End If
    Next

    Set colPingResults = Nothing
End Function


'-------------------------------------------------------------------------------------------------------------------------------------

'Loginbox subroutine - Bring up nice login box dialogue and test for login
'-------------------------------------------------------------------------------------------------------------------------------------


Function loginbox

' This function uses Internet Explorer to create a login dialog.
'
' Version:             3.11
' Last modified:       2013-11-07
'
' Arguments:           N/A
' Returns:    [array]  the user name (0) and password (1)
'                      typed in the dialog screen
'
' Written by Rob van der Woude
' http://www.robvanderwoude.com
' Error handling code written by Denis St-Pierre

    Dim objIE
    ' Create an IE object
    Set objIE = CreateObject( "InternetExplorer.Application" )
    ' specify some of the IE window's settings
    objIE.Navigate "about:blank"
    objIE.Document.title = "Laptop Login"
    objIE.ToolBar        = False
    objIE.Resizable      = False
    objIE.StatusBar      = False
    objIE.Width          = 600
    objIE.Height         = 500
    ' Center the dialog window on the screen
    With objIE.Document.parentWindow.screen
        objIE.Left = (.availWidth  - objIE.Width ) \ 2
        objIE.Top  = (.availHeight - objIE.Height) \ 2
    End With
    ' Wait till IE is ready
    Do While objIE.Busy
        WScript.Sleep 200
    Loop
    ' Insert the HTML code to prompt for user input
    objIE.Document.body.innerHTML = "<div style='padding:20px;text-align:center'><img src='http://www.uppingham.co.uk/Templates/images/print_logo.png' alt='Uppingham School' width='371' height='77' align='middle'>" & vbcrlf _
                                  & "<br><br>Pease login with your Uppingham Username and Password<table style='font-family:verdana;' cellspacing=""20""><tr nowrap>" _
                                  & "<tr></tr><td>Name:</td><td>" _
                                  & "<input type=""text"" size=""20"" " _
                                  & "autocomplete=""off"" " _
                                  & "id=""LoginName""></td></tr>" & vbcrlf _
                                  & "<tr nowrap><td>Password:</td>" _
                                  & "<td><input type=""password"" size=""20"" " _
                                  & "id=""Password""></td>" & vbcrlf _
                                  & "</tr></table>" & vbcrlf _
                                  & "<p><input type=""hidden"" id=""OK"" " _
                                  & "name=""OK"" value=""0"">" _
                                  & "<input type=""submit"" value="" OK "" " _
                                  & "onClick=""VBScript:OK.value=1""></p>" & StrLoginFailure & "</div>"
    ' Hide the scrollbars
    objIE.Document.body.style.overflow = "auto"
    ' Make the window visible
    objIE.Visible = True
    ' Set focus on input field
    objIE.Document.all.LoginName.focus

    ' Wait till the OK button has been clicked
    'On Error Resume Next
    Do While objIE.Document.all.OK.value = 0 
        WScript.Sleep 200
        If Err Then    'user clicked red X (or alt-F4) to close IE window
            WriteLog("User closed the Login Window")
            objIE.Quit
            Set objIE = Nothing
	    loginbox=false
	    Err.Clear
            Exit Function
        End if
    Loop
    
    ' Read the user input from the dialog window
    strLoginName=objIE.Document.all.LoginName.value
    strPassword=objIE.Document.all.Password.value
  
    if fnCheckAccess(strLoginName,strPassword) then
	loginbox=true
    else
	loginbox=false
    end if
     ' Close and release the object
    objIE.Quit
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' setup a nice new log file
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub InitLogFile

	'Get the computer name
	
	Dim oShell, oShellEnv,objFolder,objFile
	
	set oShell = WScript.CreateObject("WScript.Shell")
	set oShellEnv = oShell.Environment("Process")
	StrComputerName  = oShellEnv("ComputerName")


	strVLogFile = StrComputerName & " - LaptopLogin log file.txt" 

	' Check that the strVLogDir folder exists
	If objFSO.FolderExists(strVLogDir) Then
   		Set objFolder = objFSO.GetFolder(strVLogDir)
   	Else 'Create the folder if it doesn't exist
   		Set objFolder = objFSO.CreateFolder(strVLogDir)
	End If

	'Create each computer log file if required
	If not objFSO.FileExists(strVLogDir & strVLogFile) then
		Set objFile = objFSO.CreateTextFile(strVLogDir & strVLogFile)
		objFile.Close
	End if
	

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Adds a line vLineStr to the log.        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
function WriteLog(vLineStr)

Const ForAppending = 8
    
    Dim objTextFile
    Set objTextFile = objFSO.OpenTextFile _
    (strVLogDir & strVLogFile, ForAppending, True)

    objTextFile.WriteLine(Now & " - " & vLineStr)
    objTextFile.Close
    
End Function


'-------------------------------------------------------------------------------------------------------------------------------------

'Login function - validates the loginname and password and writes to a log file
'-------------------------------------------------------------------------------------------------------------------------------------


Function fnCheckAccess(strLoginName,strPassword)

Dim objDS, objDomain

 
 
    Const ADS_SECURE_AUTHENTICATION = 1
   
    Set objDS = GetObject("LDAP:")
    On Error Resume Next
    Set objDomain = objDS.OpenDSObject("LDAP://" & strDomain, strLoginName, strPassword, ADS_SECURE_AUTHENTICATION)
    If Err.Number <> 0 Then
	WriteLog("Failed to login " & strLoginName & ", Error message : " & Err.Description)  
        fnCheckAccess=false
	Err.Clear
    Else
        fnCheckAccess=true
	WriteLog("Login Success for " & strLoginName)
	
    End If

End Function


'-------------------------------------------------------------------------------------------------------------------------------------

'Login function - validates the loginname and password and writes to a log file
'-------------------------------------------------------------------------------------------------------------------------------------

Sub MapHomeDrive()
	
	Dim strUserDN,objUser,strHomePath,strHomeDrive,q
	
	Const ADS_NAME_TYPE_NT4      =   3
	Const ADS_NAME_TYPE_1779   =   1
	
	'set object translation values and then pull user dn
	
	objTrans.Set ADS_NAME_TYPE_NT4, strDomain & "\" & strLoginName
	strUserDN   =   UCase(objTrans.Get(ADS_NAME_TYPE_1779))

	 Set objUser   =   GetObject("LDAP://" & strUserDN)
	
	strHomeDrive   =   objUser.HomeDrive
	strHomePath    =   objUser.HomeDirectory
	
	'attempt to free up the drive letter prior to mapping
	
	For q = 1 To 3
		If objFSO.DriveExists(strHomeDrive) Then
			objNetwork.RemoveNetworkDrive strHomeDrive, True, True
			WScript.Sleep 1000
		Else
			Exit For
		End If
	Next
	'map drive
	If strHomePath <> "" Then 
		objNetwork.MapNetworkDrive strHomeDrive, strHomePath
		WriteLog("Mapped " &  strHomeDrive & "  to " & strHomePath & " for " & strLoginName)
	end if
	
	 If Err.Number <> 0 Then
		WriteLog("Failed to map " &  strHomeDrive & "  to " & strHomePath & " for " & strLoginName)  
		Err.Clear
        end if
	
End Sub
		
		
		
		
		
		
		
		
		
		
		
		
	