Const RETURNONLYFSDIRS = &H1 
Const NONEWFOLDERBUTTON = &H200 
  
Set oShell = CreateObject("Shell.Application") 
Set oFolder = oShell.BrowseForFolder(&H0&, "Choisir le dossier PLAT d'une affaire", RETURNONLYFSDIRS+NONEWFOLDERBUTTON, "x:\") 

If oFolder is Nothing Then  
	MsgBox "Abandon de l'opérateur",vbCritical
Else 
	Set oFolderItem = oFolder.Self 
	Racine = oFolderItem.path
	reponse = msgbox("Lancer la convertion des fichiers CN?",VbQuestion+VbYesNo, "CN to Plasma")
	If (reponse = 6) then
		Public Const tssPattern = "nc1"
		Const ForReading = 1
		Const ForWriting = 2
		Set Fso = CreateObject("Scripting.FileSystemObject")
		If Fso.FolderExists(Racine+"\plat_range") Then
			'
		Else
			Set objFolder=Fso.CreateFolder(Racine+"\plat_range")
		End If

		Set f = Fso.GetFolder(Racine+"\")
		Set colSubfolders = f.Subfolders
		set fs = Fso.GetFolder(Racine+"\")
		Set fc = fs.Files
		
		For Each f1 in fc
		
			If Split(f1.name, ".")(1) = tssPattern then 
			
				lname = (Len(f1.name)-3)
				newname = Left(f1.name, lname) + "dst"
				NewFile = Racine+"\plat_range\"+newname	
				
				If Fso.FileExists(Newfile) Then
					'
				Else 
					Set ObjFile1 = Fso.createtextFile(Newfile)  
					objFile1.Close
				End If

				Set objFile1 = Fso.OpenTextFile(NewFile, ForWriting)			
		
				File = Racine+"\"+f1.name
				Set Fso = CreateObject("Scripting.FileSystemObject" )
				Set objFile = Fso.OpenTextFile(File, ForReading)
				strText = objFile.ReadAll
				objFile.Close
				tb = split(strText,Chr(10)) 
				
				strTextNew = ""
				testKA = 0
				
				For i = LBound(tb) to UBound(tb)

					If(inStr(tb(i),"KO")) then
						testKA = 1
					End If
			
					If(inStr(tb(i),"BO") or inStr(tb(i),"IK") or inStr(tb(i),"PU") or inStr(tb(i),"KA") or inStr(tb(i),"SC") or inStr(tb(i),"TO") or inStr(tb(i),"UE") or inStr(tb(i),"PR") or inStr(tb(i),"AK") or inStr(tb(i),"EN")) then
						testKA = 0
					End If
			
					If(testKA = 0) then
						If(inStr(tb(i),"KO")) then
							objFile1.WriteLine("PU")
						Else
							objFile1.WriteLine(tb(i)) 
						End If
					End If 				
					
				Next

				objFile.Close
			End If
		Next
	End If
End If
Set oFolderItem = Nothing 
Set oFolder = Nothing 
Set oShell = Nothing

rem XS_DSTV_NO_SAWING_ANGLES_FOR_PLATES_NEEDED=FALSE
rem XS_DSTV_CREATE_AK_BLOCK_FOR_ALL_PROFILES = false