<%

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
':: File    : clsFSO.asp
':: Author  : Tom Nance (SkyDogg) - www.SkyPortal.net	
':: Date    : 07/2007
':: Subject : File System Object class
':: 		: 
'::		    : METHODS
'::			: 
'::>>		: DisplayLog(sFileURL,H,W)
'::			: Displays the file (sFileURL) in an iFrame
'::  		: sFileURL   -> Path to file from site root
'::			: H 		 -> iFrame Height.
'::			: W 		 -> iFrame Width.
'::			: 
'::>>		: WriteTextFile(strFile, strMessage)
'::			: Write plain text (strMessage) to file (strFile)
'::  		: strFile    -> Full text file path.
'::  		: 			 -> If file does not exist, the class create it.
'::			: strMessage -> Text to add to file.
'::			: 
'::>>		: AppendTextFile(strFile, strMessage)
'::			: Appends plain text (strMessage) to file (strFile)
'::  		: strFile    -> Full text file path.
'::  		: 			 -> If file does not exist, the class create it.
'::			: strMessage -> Text to add to file. One line by message
'::			: 
'::>>		: GetFileText(strFile)
'::         : Return file text content
'::			: strFile    -> Full file path 
'::			: 
'::>>		: DeleteFile(strFile)
'::			: Delete a file
'::			: strFile    -> Full file Path 
'::			: 
'::>>		: DeleteFolder(sDirectory)
'::			: Delete especific directory
'::			: sDirectory -> Full directory path 
'::			: 
'::>>		: FileExist(strFile)
'::			: Check if file exists 
'::			: strFile -> Full file path to check
'::			: 
'::>>		: FolderExist(sDirectory)
'::			: Check if directory exists
'::			: sDirectory -> Full directory path to check 
'::			: 
'::>>		: CreateFolder (sFolderPath)
'::			: Creates a folder at the path specified
'::			: If parent directory not exist, it will be created
'::			: sFolderPath -> Full directory path to create
'::			: 
'::>>		: GetFileInformation (sFilePath)
'::			: Return a FileType variable with file information
'::			: If  can't access the file, return an empty 
'::			: sFilePath -> full file path 
'::			: 
'::>>		: GetFolderInformation (sFolderPath)
'::			: Return a FileType variable with folder information
'::			: If can't access the folder, return an empty 
'::			: sFolderPath -> full directory path 
'::			: 
'::>>		: GetAllFolderInformation(sFolderPath)
'::			: Return an FileType array with all files and
'::			: directories in sFolderPath
'::			: sFolderPath -> full folder path to scan 
'::			: 
'::>>		: CopyFile(sFromPath,sToPath)
'::			: Copy a file from source directory to target directory.
'::			: If target directory not exist, it will be created
'::			: sFromPath	  	-> source file path to move
'::			: sToPath     	-> target path
'::			: 
'::>>		: MoveFile(sFromPath,sToPath)
'::			: Move a file from source directory to target directory
'::			: If target directory not exist, it will be created
'::			: sFromPath	  	-> source file path to move
'::			: sToPath     	-> target path
'::			: 
'::>>		: CopyFolder(sFromPath,sToPath)
'::			: Copy a folder from source directory to target directory
'::			: If target parent directory not exist, it will be created
'::			: sFromPath	  	-> source file path to move
'::			: sToPath     	-> target path
'::			: 
'::>>		: MoveFolder(sFromPath,sToPath)
'::			: Move a folder from source directory to target directory
'::			: If target parent directory not exist, it will be created
'::			: sFromPath	  	-> source file path to move
'::			: sToPath     	-> target path
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'::			: 
'::			: PROPERTIES
'::			: 
'::			: Version 		-> Class version 
'::			: WriteToLog 	-> True/False - Write errors to log file 
'::			: LogFile 		-> Path from root to file to write errors to 
'::			: errCount 		-> Count of errors during routine
'::			: Module 		-> Module tracking name
'::			: LogFolder 	-> Path to folder containing log file 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
':: Constants returned by File.Attributes
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const FileAttrNormal   = 0
Const FileAttrReadOnly = 1
Const FileAttrHidden = 2
Const FileAttrSystem = 4
Const FileAttrVolume = 8
Const FileAttrDirectory = 16
Const FileAttrArchive = 32 
Const FileAttrAlias = 1024
Const FileAttrCompressed = 2048

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
':: Constants for opening files
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const OpenFileForReading = 1 
Const OpenFileForWriting = 2 
Const OpenFileForAppending = 8 

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
':: DataTypes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Class clsFileType
	Public Path
	Public Name
	Public FileType
	Public Attribs
	Public Created
	Public Accessed
	Public Modified
	Public Size
End Class

Class clsSFSO

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: private variables
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Private cFSO
	Private bWriteToLog
	Private dCurTime
	Private errCnt
	Private sLogFolder
	Private sLogFile
	Private pLogFile
	Private sModule
	Private SysLogFolder
	Private SysLogFile
	Private pSysLogFile
	Private SysModule

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: Properties
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Property Get Version()
		Version = "1.0.0"
	End Property

	Public Property Get Author()
		Author = "Tom Nance (SkyDogg) - www.skyportal.net"
	End Property

	Public Property Get errCount()
		errCount = errCnt
	End Property

	Public Property Get LogFolder()
		LogFolder = sLogFolder
	End Property
	Public Property Let LogFolder(sLog)
		sLogFolder = sLog
	    pLogFile = server.MapPath(sLogFolder & sLogFile)
	End Property

	Public Property Get LogFile()
		LogFile = sLogFile
	End Property
	Public Property Let LogFile(sLog)
		sLogFile = sLog
	    pLogFile = server.MapPath(sLogFolder & sLogFile)
	End Property
	Public Property Get LogFilePath()
		LogFilePath = pLogFile
	End Property
	
	Public Property Get WriteToLog()
		WriteToLog = bWriteToLog
	End Property
	Public Property Let WriteToLog(bTF)
		bWriteToLog = bTF
	End Property
	
	Public Property Get Module()
		Module = sModule
	End Property
	Public Property Let Module(m)
		sModule = m
	End Property

	Private Sub SetAttribute(strFile,intValue)
		Dim objFile
		Set objFile = cFSO.GetFile(Cstr(strFile))
		objFile.attributes = intValue
		Set objFile = Nothing
			if err.number <> 0 then
				WriteLogFile("SetAttribute: " & err.Description)
			end if	
	End Sub


	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: MoveFile
	':: Return "" if ok else return error message
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Function MoveFile(sFromPath,sToPath)
		Dim sReturn 
		Dim sToFolder
		Dim sToFile
		Dim fo
		sReturn = ""

		On error resume next 
  	    err.clear
		sToFolder = left(sToPath,inStrRev(sToPath,"\")-1)
		sToFile   = mid(sToPath,inStrRev(sToPath,"\")+1,len(sToPath))
		
		'sFromPath = sFromPath & "\" & sOrigFile 
		'sToPath= sToPath & "\" & sToFile
		if FileExist(sFromPath) then 
			Call SetAttribute(sFromPath,0)		
			'Check if directory exists
			If Not FolderExist(sToFolder) then
				sReturn = CreateFolder(sToFolder)
				If not sReturn = "" then 
				  if bWriteToLog and sReturn <> "" then
				    sReturn = sReturn & WriteLogFile("MoveFile1" & sReturn)
					errCnt = errCnt + 1
				  end if
				  MoveFile = sReturn
				  Exit Function 
				end if
			End if 
			set fo = cFSO.GetFile(sFromPath)
			if err.number <> 0 then
				WriteLogFile("MoveFile: cFSO.GetFolder: " & err.Description & " : " & sFromPath)
			end if
			fo.Copy(sToPath)
			if err.number <> 0 then 
				sReturn = err.Description
				WriteLogFile("MoveFile: fo.Copy: " & err.Description & " : " & sToPath)
			else
				'Check copy
				If FileExist(sToPath) then 
					'delete source file
					sReturn = DeleteFile(sFromPath) = ""
					'DeleteFile(sFromPath)
				else
					sReturn = "Can not move file " & sFromPath
				End if  
			End If 
			set fo = nothing
		else
	        sReturn = "Can not move file - File does not exist" & " " & sFromPath
		end if
		if bWriteToLog and len(sReturn) > 5 then
		    sReturn = sReturn & WriteLogFile("MoveFile2: " & sReturn)
			errCnt = errCnt + 1
		end if
		MoveFile = sReturn
	End Function


	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: MoveFolder
	':: Return "" if ok else return error message
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Function MoveFolder(sFromPath,sToPath)
		Dim sReturn 
		Dim sToParent
		Dim fo
		sReturn = ""

		On error resume next 
  	    err.clear
		sToParent = left(sToPath,inStrRev(sToPath,"\")-1)
		
		if FolderExist(sFromPath) then 
			Call SetAttribute(sFromPath,0)		
			'Check if directory exists
			If Not FolderExist(sToParent) then
				sReturn = CreateFolder(sToParent)
				If not sReturn = "" then 
				  if bWriteToLog and sReturn <> "" then
				    sReturn = sReturn & WriteLogFile("MoveFile1" & sReturn)
					errCnt = errCnt + 1
				  end if
				  MoveFolder = sReturn
				  Exit Function 
				end if
			End if 
			set fo = cFSO.GetFolder(sFromPath)
			fo.Copy(sToPath)
			if err.number <> 0 then 
				sReturn = err.Description
			else
				'Check copy
				If FolderExist(sToPath) then 
					'delete source file
					sReturn = DeleteFolder(sFromPath)
				else
					sReturn = "Can not move folder " & sFromPath
				End if
			End If 
			set fo = nothing
		else
	        sReturn = "Can not move folder - Folder does not exist" & " " & sFromPath
		end if
		if bWriteToLog and sReturn <> "" then
		    sReturn = sReturn & WriteLogFile("MoveFile2" & sReturn)
			errCnt = errCnt + 1
		end if
		MoveFolder = sReturn
	End Function


	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: CopyFile
	':: Return "" if ok else return error message
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Function CopyFile(sFromPath,sToPath)
		Dim sReturn 
		Dim sToFolder
		Dim sToFile
		Dim fo
		sReturn = ""

		On error resume next 
  	    err.clear
		sToFolder = left(sToPath,inStrRev(sToPath,"\")-1)
		sToFile   = mid(sToPath,inStrRev(sToPath,"\")+1,len(sToPath))
		
		'sFromPath = sFromPath & "\" & sOrigFile 
		'sToPath= sToPath & "\" & sToFile
		if FileExist(sFromPath) then 
			Call SetAttribute(sFromPath,0)		
			'Check if directory exists
			If Not FolderExist(sToFolder) then
				sReturn = CreateFolder(sToFolder)
				If not sReturn = "" then 
				  if bWriteToLog and sReturn <> "" then
				    sReturn = sReturn & WriteLogFile(sReturn)
					errCnt = errCnt + 1
				  end if
				  CopyFile = sReturn
				  Exit Function 
				end if
			End if 
			set fo = cFSO.GetFile(sFromPath)
			fo.Copy(sToPath)
			if err.number <> 0 then 
				sReturn = err.Description
			else
				'Check copy
				If FileExist(sToPath) then 
					'success
				else
					sReturn = "Can not copy file " & sFromPath
				End if  
			End If 
			set fo = nothing
		else
	        sReturn = "Can not copy file - File does not exist" & " " & sFromPath
		end if
		if bWriteToLog and sReturn <> "" then
		    sReturn = sReturn & WriteLogFile(sReturn)
			errCnt = errCnt + 1
		end if
		CopyFile = sReturn
	End Function


	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: CopyFolder
	':: Return "" if ok else return error message
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Function CopyFolder(sFromPath,sToPath)
		Dim sReturn 
		Dim sToParent
		Dim fo
		sReturn = ""

		On error resume next 
  	    err.clear
		sToParent = left(sToPath,inStrRev(sToPath,"\")-1)
		
		if FolderExist(sFromPath) then 
			Call SetAttribute(sFromPath,0)		
			'Check if directory exists
			If Not FolderExist(sToParent) then
				sReturn = CreateFolder(sToParent)
				If not sReturn = "" then 
				  if bWriteToLog and sReturn <> "" then
				    sReturn = sReturn & WriteLogFile(sReturn)
					errCnt = errCnt + 1
				  end if
				  CopyFolder = sReturn
				  Exit Function 
				end if
			End if 
			set fo = cFSO.GetFolder(sFromPath)
			fo.Copy(sToPath)
			if err.number <> 0 then 
				sReturn = err.Description
			else
				'Check copy
				If FolderExist(sToPath) then 
					'success
				else
					sReturn = "Can not copy folder " & sFromPath
				End if
			End If 
			set fo = nothing
		else
	        sReturn = "Can not copy folder - Folder does not exist" & " " & sFromPath
		end if
		if bWriteToLog and sReturn <> "" then
		    sReturn = sReturn & WriteLogFile(sReturn)
			errCnt = errCnt + 1
		end if
		CopyFolder = sReturn
	End Function

	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: CreateTextFile
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Function CreateTextFile(strFile)
		Dim TextStream
		Dim sFile 
		Dim sFolder
		Dim Folder 
	    Dim sReturn
	    sReturn = ""
		sFolder = left(strFile,inStrRev(strFile,"\")-1)
		sFile   = mid(strFile,inStrRev(strFile,"\")+1,len(strFile))
		if not FolderExist(sFolder) then
		  CreateFolder(sFolder)
		end if
		Set Folder = cFSO.GetFolder(sFolder)
		Set TextStream = Folder.CreateTextFile(sFile)
	    if err.number <> 0 then 
			sReturn = err.Description 
			errCnt = errCnt + 1
	    end if
		TextStream.Close
		CreateTextFile = sReturn
	End Function
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: AppendTextFile
	':: Return "" if ok else return error message
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Function AppendTextFile(strFile, strMessage)
	  if strMessage <> "" then
	   Dim TextStream
	   Dim sReturn
	   sReturn = ""
	   
	   On Error Resume Next
	   If not FileExist(strFile) then 
		CreateTextFile(strFile)
	   End If 
	   Set TextStream = cFSO.OpenTextFile(strFile, OpenFileForAppending) 
	   if err.number <> 0 then 
			sReturn = err.Description 
			errCnt = errCnt + 1
	   else
			TextStream.Write(strMessage)
			TextStream.WriteBlankLines(1)
			TextStream.Close
	   End if
	  else
	     sReturn = ""
	  end if
	   AppendTextFile = sReturn
	End Function
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: WriteLogFile
	':: Return "" if ok else return error message
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Function WriteLogFile(sTxt)
	   Dim TextStream
	   Dim sReturn
	   sReturn = ""
	   On Error Resume Next
	   If not FileExist(pLogFile) then 
		 CreateTextFile(pLogFile)
		 if sLogFile = SysLogFile then
		   WriteTextFile pLogFile,SysModule
		 else
		   if sModule = "" then
		     WriteTextFile pLogFile,"No Name Log"
		   else
		     WriteTextFile pLogFile,sModule
		   end if
		 end if
		 AppendTextFile pLogFile,"Created: " & Date()
		 AppendTextFile pLogFile,"[]"
	   End If 
	   if sTxt <> "[]" then
	     if sLogFile = SysLogFile and sModule <> "" then
	       sTxt = dCurTime & " [" & sModule & "] (" & sScript & ")" & sTxt
		 else
	       sTxt = dCurTime & " (" & sScript & ")" & sTxt
		 end if
	   end if
	   Set TextStream = cFSO.OpenTextFile(pLogFile, OpenFileForAppending) 
	   if err.number <> 0 then 
			sReturn = err.Description 
			errCnt = errCnt + 1
	   else
			TextStream.Write(sTxt)
			TextStream.WriteBlankLines(1)
			TextStream.Close
	   End if
	   WriteLogFile = sReturn
	End Function
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: WriteTextFile
	':: Return "" if ok else return error message
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Function WriteTextFile(strFile, strMessage)
	  if strMessage <> "" then
	   Dim TextStream
	   Dim sReturn
	   
	   On Error Resume Next
	   sReturn = ""
	   If not FileExist(strFile) then 
		CreateTextFile(strFile)
	   End If 
	   Set TextStream = cFSO.OpenTextFile(strFile, OpenFileForWriting) 
	   if err.number <> 0 then 
			sReturn = err.Description 
			errCnt = errCnt + 1
	   else
			TextStream.Write(strMessage)
			TextStream.WriteBlankLines(1)
			TextStream.Close
	   End if
	  else
	     sReturn = ""
	  end if
	   WriteTextFile = sReturn
	End Function
	
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: GetFileText
	':: Return file content if ok else return error message
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Function GetFileText(strFile)
	   Dim TextStream
	   Dim sReturn 
	   Dim File

	   On Error Resume Next
	   sReturn  = ""
	   If not FileExist(strFile) then exit function
	   Set TextStream = cFSO.OpenTextFile(strFile, OpenFileForReading)
	   If err.number <> 0 then  
			sReturn  = err.Description 
			errCnt = errCnt + 1
	   else
			sReturn  = TextStream.ReadAll 
	   End if 
	   TextStream.Close
       GetFileText = sReturn 

	   'OTHER WAY 
	   'Set File = cFSO.GetFile(strFile)
	   'Set TextStream = File.OpenAsTextStream(OpenFileForReading)
	   'Do While Not TextStream.AtEndOfStream
	   '   sReturn  = sReturn  & TextStream.ReadLine & vbLfCr
	   'Loop
	   
	End Function


	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: DeleteFile
	':: Return "" if ok, else return error message 
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Function DeleteFile(strFile)
	  Dim sReturn
	  Dim tf
	  sReturn = ""
  	    on error resume next
  	    err.clear
  		if cFSO.FileExists(strFile) then
		  set tf = cFSO.GetFile(strFile)
		  tf.Delete
		  if err.number <> 0 then 
	        sReturn = Err.Description
			'WriteLogFile(err.Description)
		  end if
		  set tf = nothing
		else
	      sReturn = "Cannot delete file - File does not exist" & " " & strFile
  		end if
  	    on error goto 0
		if bWriteToLog and sReturn <> "" then
		    'sReturn = sReturn & WriteLogFile("DeleteFile: " & sReturn & " : " & strFile)
			WriteLogFile("DeleteFile: " & sReturn & " : " & strFile)
			errCnt = errCnt + 1
		end if
	  DeleteFile = sReturn 
	End Function

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: DeleteDirectory
	':: Return "" if ok, else return error message 
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Function DeleteFolder(sDirectory)
	  Dim sReturn
	  Dim tf
	  sReturn = ""
  	    on error resume next
  	    Err.clear
  		if cFSO.FolderExists(sDirectory) then
		  set tf = cFSO.GetFolder(sDirectory)
		  tf.Delete
		  if err.number <> 0 then
	        sReturn = Err.Description 
		  end if
		  set tf = nothing
		else
		  sReturn = "Cannot delete folder"
  		end if
  	    on error goto 0
		if bWriteToLog and sReturn <> "" then
		    sReturn = sReturn & WriteLogFile(sReturn)
			WriteLogFile("DeleteFolder: "& sReturn & " : "& sDirectory)
			errCnt = errCnt + 1
		end if
	  DeleteFolder = sReturn 
	End Function

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: FileExist
	':: Return true if file exists, otherwise false
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Function FileExist(strFile)
			err.clear()
		FileExist = cFSO.FileExists(strFile) 
			if err.number <> 0 then
				WriteLogFile("FileExist:" & strFile & "-- " & err.Description)
			    err.clear()
			end if
			'err.number = 0
    End Function 

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: FolderExist
	':: return true if directory exists, otherwise false
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Function FolderExist(sDirectory)
			err.clear()
		FolderExist = cFSO.FolderExists(sDirectory) 
			if err.number <> 0 then
				WriteLogFile("FolderExist: " & err.Description)
			err.clear()
			end if
    End Function 
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: CreateDirectory
	':: Return "" if ok, else return error message 
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Function CreateFolder(sFolderPath)
	  Dim bCreate
	  bCreate = ""
		On Error Resume Next
		If Not FolderExist(sFolderPath) Then
			bCreate = CreateFolder(cFSO.GetParentFolderName(sFolderPath))
		    IF bCreate = "" then 
				cFSO.CreateFolder sFolderPath
				bCreate = err.Description
			End if 
		End if 
		if bWriteToLog and bCreate <> "" then
		    bCreate = bCreate & WriteLogFile("CreateFolder: " & bCreate)
			errCnt = errCnt + 1
		end if
		on error goto 0
	  CreateFolder = bCreate
	End Function 

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: ShowFileAttr
	':: Return an string with file attributes 
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Private Function ShowFileAttr(File) ' File can be a file or folder
	   Dim S
	   Dim Attr
	   Attr = File.Attributes

	   If Attr = 0 Then
	      ShowFileAttr = "Normal"
	      Exit Function
	   End If

	   If Attr And FileAttrDirectory Then S = S & "Directory "
	   If Attr And FileAttrReadOnly Then S = S & "Read-Only "
	   If Attr And FileAttrHidden Then S = S & "Hidden "
	   If Attr And FileAttrSystem Then S = S & "System "
	   If Attr And FileAttrVolume Then S = S & "Volume "
	   If Attr And FileAttrArchive Then S = S & "Archive "
	   If Attr And FileAttrAlias Then S = S & "Alias "
	   If Attr And FileAttrCompressed Then S = S & "Compressed "

	  ShowFileAttr = S
	End Function

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: GetFileInformation
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Function GetFileInformation(sFilePath)
	   Dim File
	   Set GetFileInformation = new clsFileType
	   If Not cFSO.FileExists(sFilePath) Then Exit Function
	   Set File = cFSO.GetFile(sFilePath)
	   Set GetFileInformation = GenerateFileInformation(File)
	End Function 
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: GenerateFileInformation
	':: Return a FileType variable with file information
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Private Function GenerateFileInformation(File)
	   Dim S
	   Set S = new clsFileType
	   
	   S.Path     = File.Path
	   S.Name     = File.Name
	   S.FileType = File.Type
	   S.Attribs  = ShowFileAttr(File)
	   S.Created  = File.DateCreated
	   S.Accessed = File.DateLastAccessed
	   S.Modified = File.DateLastModified
	   S.Size     = File.Size
		
	   Set GenerateFileInformation = S
	End Function

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: GetFolderInformation
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Function GetFolderInformation(sFolderPath)
	   Dim Folder
	   Set GetFolderInformation = new clsFileType
	   If Not cFSO.FolderExists(sFolderPath) Then Exit Function
	   Set Folder = cFSO.GetFolder(sFolderPath)
	   Set GetFolderInformation = GenerateFolderInformation(Folder)
	End Function 
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: GenerateFolderInformation
	':: Return a clsFileType type variable with folder properties
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Private Function GenerateFolderInformation(Folder)
	   Dim S
	   On Error Resume Next
	   Set S = new clsFileType
	   S.Path     = Folder.Path
	   S.Name     = Folder.Name
	   S.FileType = "Directory"
	   S.Attribs  = ShowFileAttr(Folder)
	   S.Created  = Folder.DateCreated
	   S.Accessed = Folder.DateLastAccessed
	   S.Modified = Folder.DateLastModified
	   S.Size	  = Folder.Size
	   Set GenerateFolderInformation = S
	End Function

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: GetAllFolderInformation
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Function GetAllFolderInformation(sFolderPath)
	   Dim Folder
	   If Not cFSO.FolderExists(sFolderPath) Then Exit Function
	   Set Folder = cFSO.GetFolder(sFolderPath)
	   GetAllFolderInformation = GenerateAllFolderInformation(Folder)
	End Function

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: GenerateAllFolderInformation
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Private Function GenerateAllFolderInformation(Folder)
	   Dim S
	   Dim SubFolders
	   Dim SubFolder
	   Dim Files
	   Dim File
	   Dim arrInformation()
	   Dim Index 

	   Set Files = Folder.Files
	   Index = 0
	   If Files.Count <> 0 Then
	      For Each File In Files
			Redim preserve arrInformation(Index)
	        Set arrInformation(Index) = GenerateFileInformation(File)
	        Index = Index + 1
	      Next
	   End If
	   Set SubFolders = Folder.SubFolders
	   If SubFolders.Count <> 0 Then
	      For Each SubFolder In SubFolders
			Redim preserve arrInformation(Index)
	        Set arrInformation(Index) = GenerateFolderInformation(SubFolder)
	        Index = Index + 1
	      Next
	   End If
	   GenerateAllFolderInformation = arrInformation
	End Function

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: DisplayLog
	':: Displays the file (sFileURL) in an iFrame
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Sub DisplayLog(sFileURL,H,W)
	  if FileExist(server.MapPath(sFileURL)) then
	    Response.Write "<iframe src=""" & sFileURL & """ frameborder=""0"" height=""" & H & """ width=""" & W & """ id=""tabiframe"" class=""tabiframe"" scrolling=""auto""></iframe>"
	  else
	    Response.Write "File not found" & ": " & sFileURL
	  end if
	End Sub


	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: BuildSysLogName
	':: Builds the days System log file name
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Private Function BuildSysLogName()
	  Dim sReturn, cY, cM, cD
	  sReturn = "sys_log_"
	  cY = Year(date)
	  cM = dNum(Month(date))
	  cD = dNum(Day(date))
	  sReturn = sReturn & cY & cM & cD & ".txt"
	  BuildSysLogName = sReturn
	End Function
	
	Private Function dNum(n)
	  Dim sReturn
	  sReturn = ""
	  If len(n) = 1 then
	    sReturn = "0" & n
	  else
	    sReturn = n
	  end if
	  dNum = sReturn
	End Function
	

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: Reset class variables
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Sub fsoReset()
	  sLogFolder = SysLogFolder
	  sLogFile = SysLogFile
	  pLogFile = server.MapPath(sLogFolder & sLogFile)
	  sModule = ""
	  bWriteToLog = true
	  dCurTime = now() & " "
	end Sub

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	':: Initialize and Terminate clase
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Private Sub Class_Initialize()
	   'Create FileSystemObject
	   Set cFSO = Server.CreateObject("Scripting.FileSystemObject")
			if err.number <> 0 then 
				sReturn = err.Description
				WriteLogFile("Class_Initialize: " & sReturn)
			end if
	   errCnt = 0
	   SysLogFolder = "files/sp_Logs/"
	   SysModule = "SkyPortal System Log"
	   SysLogFile = BuildSysLogName
	   pSysLogFile = server.MapPath(SysLogFolder & SysLogFile)
	   fsoReset()
	End Sub
	  
	Private Sub Class_Terminate()
		if not errCnt = 0 then
		  'AppendTextFile pLogFile,"[]"
		end if
		'Destroy cFSO object 
		Set cFSO = nothing
	End Sub
	 
	
End Class

%>
