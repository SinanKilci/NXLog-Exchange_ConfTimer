Dim strNow, strDD, strMM, strYYYY, strFulldate
strYYYY = DatePart("yyyy",Now())
strMM = Right("0" & DatePart("m",Now()),2)
strDD = Right("0" & DatePart("d",Now()),2)
fulldate = strYYYY & strMM & strDD
Const ForReading = 1
Const ForWriting = 2
Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Pattern = "(\s+File\s+.)(C:\\\\Program Files\\\\Microsoft\\\\Exchange Server\\\\V15\\\\TransportRoles\\\\Logs\\\\MessageTracking\\\\MSGTRK)(\w{0,2})(\d{8})(\*-\*\.LOG)(.)"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Program Files (x86)\nxlog\conf\nxlog.conf", ForReading)
Do Until objFile.AtEndOfStream
  strSearchString = objFile.ReadLine
  Set colMatches = objRegEx.Execute(strSearchString)
  If colMatches.Count > 0 Then
    For Each strMatch in colMatches
		  strIdentifier = objRegEx.Replace(strSearchString,"$3")
		  strSearchStr = objRegEx.Replace(strSearchString,"$2$3$4$5")
		  If strIdentifier="MD" Then
			  strNewStringMD = objRegEx.Replace(strSearchString,"C:\\Program Files\\Microsoft\\Exchange Server\\V15\\TransportRoles\\Logs\\MessageTracking\\MSGTRKMD"&fulldate&"*-*.LOG")
			  strSearchStrMD = strSearchStr
		  ElseIf strIdentifier="MS" Then
			  strNewStringMS = objRegEx.Replace(strSearchString,"C:\\Program Files\\Microsoft\\Exchange Server\\V15\\TransportRoles\\Logs\\MessageTracking\\MSGTRKMS"&fulldate&"*-*.LOG")
			  strSearchStrMS = strSearchStr
		  ElseIf strIdentifier = "" Then
			  strNewString = objRegEx.Replace(strSearchString,"C:\\Program Files\\Microsoft\\Exchange Server\\V15\\TransportRoles\\Logs\\MessageTracking\\MSGTRK"&fulldate&"*-*.LOG")
			  strSearchStrE = strSearchStr
		  End If
    Next
  End If
Loop
objFile.Close

Set objFile = objFSO.OpenTextFile("C:\Program Files (x86)\nxlog\conf\nxlog.conf", ForReading)
strText = objFile.ReadAll
objFile.Close
strNewText = Replace(strText, strSearchStrE, strNewString)
Set objFile = objFSO.OpenTextFile("C:\Program Files (x86)\nxlog\conf\nxlog.conf", ForWriting)
objFile.Write strNewText
objFile.Close


Set objFile = objFSO.OpenTextFile("C:\Program Files (x86)\nxlog\conf\nxlog.conf", ForReading)
strText = objFile.ReadAll
objFile.Close
strNewTextMS = Replace(strText, strSearchStrMS, strNewStringMS)
Set objFile = objFSO.OpenTextFile("C:\Program Files (x86)\nxlog\conf\nxlog.conf", ForWriting)
objFile.Write strNewTextMS
objFile.Close


Set objFile = objFSO.OpenTextFile("C:\Program Files (x86)\nxlog\conf\nxlog.conf", ForReading)
strText = objFile.ReadAll
objFile.Close
strNewTextMD = Replace(strText, strSearchStrMD, strNewStringMD)
Set objFile = objFSO.OpenTextFile("C:\Program Files (x86)\nxlog\conf\nxlog.conf", ForWriting)
objFile.Write strNewTextMD
objFile.Close

strServiceName = "nxlog"
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colListOfServices = objWMIService.ExecQuery("Select * from Win32_Service Where Name ='" & strServiceName & "'")
For Each objService in colListOfServices
objService.StopService()
Next

strServiceName = "nxlog"
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colListOfServices = objWMIService.ExecQuery ("Select * from Win32_Service Where Name ='" & strServiceName & "'")
For Each objService in colListOfServices
objService.StartService()
Next
