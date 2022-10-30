Const RETURNONLYFSDIRS = &H1 
Const NONEWFOLDERBUTTON = &H200 
  
Set oShell = CreateObject("Shell.Application") 
Set oFolder = oShell.BrowseForFolder(&H0&, "Choisir le dossier PLAT d'une affaire", RETURNONLYFSDIRS+NONEWFOLDERBUTTON, "x:\") 

If oFolder is Nothing Then  
	MsgBox "Abandon operateur",vbCritical
Else 
	Set oFolderItem = oFolder.Self 
	Racine = oFolderItem.path
	reponse = msgbox("Lancer la convertion des fichiers CN?",VbQuestion+VbYesNo, "CN to Plasma")
	If (reponse = 6) Then
		Public Const tssPattern = "nc1"
		Public Const tssPatternnew = "nc1"
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
		Set fs = Fso.GetFolder(Racine+"\")
		Set fc = fs.Files
		
		For Each f1 in fc
			If Split(f1.name, ".")(1) = tssPattern Then 
				File = Racine+"\"+f1.name
				Set Fso = CreateObject("Scripting.FileSystemObject" )
				Set objFile = Fso.OpenTextFile(File, ForReading)
				strText = objFile.ReadAll
				objFile.Close
				tb = split(strText,Chr(10)) 
				
				test = "N"
				strTextNew = ""
				testKA = 0
				
				For i = LBound(tb) to UBound(tb)
					If(inStr(tb(i),"BO")) Then
						test = "P"
					End If

					If(inStr(tb(i),"KA")) Then
					testKA = 1
					End If
					
					If(inStr(tb(i),"IK") or inStr(tb(i),"PU") or inStr(tb(i),"KA") or inStr(tb(i),"SC") or inStr(tb(i),"TO") or inStr(tb(i),"UE") or inStr(tb(i),"PR") or inStr(tb(i),"EN")) Then
					testKA = 0
					End If
					
					If(testKA = 0) Then
						If(inStr(tb(i),"KO")) Then
							strTextNew = strTextNew + "PU" + Chr(10)
						Else
							strTextNew = strTextNew + tb(i) 
						End If
					End If 				
				next
				
				tabep = Split(tb(13), ".")
				ep = Replace(tabep(0)," ","") + "mm"
				lname = (Len(f1.name)-3)
				newname = Left(f1.name, lname) + tssPatternnew
				
				If Fso.FolderExists(Racine+"\plat_range\"+ep) Then
					'
				Else
					Set objFolder=Fso.CreateFolder(Racine+"\plat_range\"+ep)
				End If
				
				NewFile = Racine+"\plat_range\"+ep+"\"+newname

				If Fso.FileExists(Newfile) Then
					'
				Else 
					Set ObjFile = Fso.createtextFile(Newfile)  
					objFile.Close
				End If

				Set objFile = Fso.OpenTextFile(NewFile, ForWriting)
				objFile.WriteLine strTextNew
				objFile.Close
			End If

		next
	End If
End If

Set oFolderItem = Nothing 
Set oFolder = Nothing 
Set oShell = Nothing

rem XS_DSTV_NO_SAWING_ANGLES_FOR_PLATES_NEEDED=FALSE
rem XS_DSTV_CREATE_AK_BLOCK_FOR_ALL_PROFILES = false