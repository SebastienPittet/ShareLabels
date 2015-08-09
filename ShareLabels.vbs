' ¨°º©o¿,,¿o©º°¨¨°º©o¿,,¿o©º°¨¨°º©o¿,,¿o©º°¨¨°º©o¿,,¿o©º°¨

'    Author: sebastien at pittet dot org
'    Date  : 30.10.2003
'    Goal  : Label all network mappings
'  Version : ShareLabels.vbs v.1.1 (TextFile)

' ¨°º©o¿,,¿o©º°¨¨°º©o¿,,¿o©º°¨¨°º©o¿,,¿o©º°¨¨°º©o¿,,¿o©º°¨

'@@@@@@@@@@@@@
' MAIN PROGRAM
'@@@@@@@@@@@@@

'Constants & Variables

Const   Title = "ShareLabels v.1.1" 'Titre InputBoxes
Const   Share = 0 'Column Number to retrieve ShareName
Const  French = 1 'Column Number to retrieve French label
Const  German = 2 'Column Number to retrieve German label
Const English = 3 'Column Number to retrieve English label
Const ForReading = 1 'Reads from the text file


Dim strNoParams 'Text displayed if no parameters is given

strNoParams = "This Script reads a ASCII text file and set " & _
              "the labels of the network drives." & vbCrLf & vbCrLf & _
              "Type the path of the file containing the information the script needs." & vbCrLf & vbCrLf & _
              "Please, visit http://www.pittet.org to get informations"

'Get the command line args
  SET Parameters = Wscript.arguments

'If no command line arguments provided, prompt for file
  If Parameters.Count = 0 Then
    TextFile = InputBox(strNoParams,Title, GetThisFolderPath & "\shareLabels.txt")
  Else
    TextFile = Parameters.item(0) 'Else file containing users is the first argument
  End If

  If TextFile = "" Then
     Error=MsgBox("No input file provided. Stopping the script now.",vbokonly, Title)
     WScript.Quit(1)
  End If

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set ShareDescriptionsFile = objFSO.OpenTextFile(TextFile, ForReading)

Do While Not ShareDescriptionsFile.AtEndOfStream
    TextFileLine = ShareDescriptionsFile.ReadLine 'Read a line
    
    If Left(TextFileLine,2) = "\\" Then 'Ignore les commentaires
		MyLine = Split(TextFileLine,";") 'Récupère la ligne dans un tableau
		
	    Select Case GetOSLanguage
	      Case "French"
		      Server_n_Share = Replace(Replace(MyLine(Share), "\", "#"),"%USERNAME%", GetUserName)
		      French_Label = Replace(MyLine(French), "%USERNAME%", GetUserName)
		      Call RegCreateMountPoints(Server_n_Share, French_Label)
		      
	      Case "German"
		      Server_n_Share = Replace(Replace(MyLine(Share), "\", "#"),"%USERNAME%", GetUserName)
		      German_Label = Replace(MyLine(German), "%USERNAME%", GetUserName)
		      Call RegCreateMountPoints(Server_n_Share, German_Label)
	      
	      Case "Else"
		      Server_n_Share = Replace(Replace(MyLine(Share), "\", "#"),"%USERNAME%", GetUserName)
		      English_Label = Replace(MyLine(English), "%USERNAME%", GetUserName)
		      Call RegCreateMountPoints(Server_n_Share, English_Label)  
	    End Select   
    End if 
Loop
ShareDescriptionsFile.close
WScript.Quit(1) 

'@@@@@@@@@@@@@@@
'Functions & Sub
'@@@@@@@@@@@@@@@
'--------------------------------------------------------------------

'Retourne le chemin du dossier d'où le script est exécuté
   Function GetThisFolderPath()
      Set fso = CreateObject("Scripting.FileSystemObject")
      Set file = fso.GetFile(wscript.scriptfullname)
      GetThisFolderPath=File.ParentFolder
   End Function
'--------------------------------------------------------------------

'Création de la clé Reg pour fixer le label correspondant au share
   Sub RegCreateMountPoints(Share,Label)
      REGKEY ="HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2\" & Share & "\" & "_LabelFromDesktopINI"
      Call RegWrite(REGKEY,Label)
      Call RegWrite("HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2\" & Share & "\" & "_CommentFromDesktopINI", "")
      Call RegWrite("HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2\" & Share & "\" & "BaseClass", "Drive")
   End Sub
'--------------------------------------------------------------------

'Returns the OS Language
   Function GetOSLanguage
	Dim languageNR 'String, contains the language number
	
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colOperatingSystems = objWMIService.ExecQuery ("SELECT OSLanguage FROM Win32_OperatingSystem")
	
	For Each objOperatingSystem in colOperatingSystems
	  LanguageNR = objOperatingSystem.oslanguage
	Next
	
	Select Case languageNR
	  Case 1036
	    GetOSLanguage = "French"
	  Case 1031
	    GetOSLanguage = "German"
	  Case 0409 
	    GetOSLanguage = "English"
	  Case Else
	    GetOSLanguage = "Error"
	End Select
End Function
'--------------------------------------------------------------------
'Write a Registry Key
   Sub RegWrite(RegKey, Value)
      Dim WshShell
      Set WshShell = WScript.CreateObject("WScript.Shell")
      WshShell.RegWrite RegKey, Value
   End Sub
'--------------------------------------------------------------------

   Function GetUserName
      Set WshNetwork = WScript.CreateObject("WScript.Network")
      GetUserName = WshNetwork.UserName
   End Function
'--------------------------------------------------------------------

