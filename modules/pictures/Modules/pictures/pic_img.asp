<%
'  For examples, documentation, and your own free copy, go to:
'  http://www.freeaspupload.net
'  Note: You can copy and use this script for free and you can make changes
'  to the code, but you cannot remove the above comment.

Class FreeASPUpload
	Public UploadedFiles
	Public FormElements
	Public Errors

	Private VarArrayBinRequest
	Private StreamRequest
	Private uploadedYet

	Private Sub Class_Initialize()
		Set UploadedFiles = Server.CreateObject("Scripting.Dictionary")
		Set FormElements = Server.CreateObject("Scripting.Dictionary")
		Set Errors =  Server.CreateObject("Scripting.Dictionary")
		
		Set StreamRequest = Server.CreateObject("ADODB.Stream")
		StreamRequest.Type = 1 'adTypeBinary
		StreamRequest.Open
		uploadedYet = false
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(UploadedFiles) Then
			UploadedFiles.RemoveAll()
			Set UploadedFiles = Nothing
		End If
		If IsObject(FormElements) Then
			FormElements.RemoveAll()
			Set FormElements = Nothing
		End If
		If IsObject(Errors) Then
			Errors.RemoveAll()
			Set Errors = Nothing
		End If		
		StreamRequest.Close
		Set StreamRequest = Nothing
	End Sub

	Public Property Get Form(sIndex)
		Form = ""
		If FormElements.Exists(LCase(sIndex)) Then Form = FormElements.Item(LCase(sIndex))
	End Property

	Public Property Get Files()
		Files = UploadedFiles.Items
	End Property

	Public Property Get CriticalErrors()
		CriticalErrors = Errors.Items
	End Property
	
	'Calls Upload to extract the data from the binary request and then saves the uploaded files
	Public Sub Save(path)
		Dim streamFile, fileItem

		if Right(path, 1) <> "\" then path = path & "\"

		if not uploadedYet then Upload

		For Each fileItem In UploadedFiles.Items
			Set streamFile = Server.CreateObject("ADODB.Stream")
			streamFile.Type = 1
			streamFile.Open
			StreamRequest.Position=fileItem.Start
			StreamRequest.CopyTo streamFile, fileItem.Length
			On Error Resume Next
			streamFile.SaveToFile path & fileItem.FileName, 2
			if Err.Number <> 0 then 
				Errors.Add Err.Number, "You propably don't have the proper rights to upload here."
				Exit Sub
			end if
			on error goto 0 'reset error handling
			streamFile.close
			Set streamFile = Nothing
			fileItem.Path = path & fileItem.FileName
		 Next
	End Sub

	Public Function SaveBinRequest(path) ' For debugging purposes
		StreamRequest.SaveToFile path & "\debugStream.bin", 2
	End Function

	Public Sub DumpData() 'only works if files are plain text
		Dim i, aKeys, f
		response.write "Form Items:<br>"
		aKeys = FormElements.Keys
		For i = 0 To FormElements.Count -1 ' Iterate the array
			response.write aKeys(i) & " = " & FormElements.Item(aKeys(i)) & "<BR>"
		Next
		response.write "Uploaded Files:<br>"
		For Each f In UploadedFiles.Items
			response.write "Name: " & f.FileName & "<br>"
			response.write "Type: " & f.ContentType & "<br>"
			response.write "Start: " & f.Start & "<br>"
			response.write "Size: " & f.Length & "<br>"
		 Next
   	End Sub

	Private Sub Upload()
		Dim nCurPos, nDataBoundPos, nLastSepPos
		Dim nPosFile, nPosBound
		Dim sFieldName, osPathSep, auxStr

		'RFC1867 Tokens
		Dim vDataSep
		Dim tNewLine, tDoubleQuotes, tTerm, tFilename, tName, tContentDisp, tContentType
		tNewLine = Byte2String(Chr(13))
		tDoubleQuotes = Byte2String(Chr(34))
		tTerm = Byte2String("--")
		tFilename = Byte2String("filename=""")
		tName = Byte2String("name=""")
		tContentDisp = Byte2String("Content-Disposition")
		tContentType = Byte2String("Content-Type:")

		uploadedYet = true

		on error resume next
		VarArrayBinRequest = Request.BinaryRead(Request.TotalBytes)
		if Err.Number <> 0 then 
			Errors.Add Err.Number, "The most likely cause for this error is the incorrect setup of AspMaxRequestEntityAllowed in IIS MetaBase. Please see instructions in the <A HREF='http://www.freeaspupload.net/freeaspupload/requirements.asp'>requirements page of freeaspupload.net</A>.<p>"
			Exit Sub
		end if
		on error goto 0 'reset error handling

		nCurPos = FindToken(tNewLine,1) 'Note: nCurPos is 1-based (and so is InstrB, MidB, etc)

		If nCurPos <= 1  Then Exit Sub
		 
		'vDataSep is a separator like -----------------------------21763138716045
		vDataSep = MidB(VarArrayBinRequest, 1, nCurPos-1)

		'Start of current separator
		nDataBoundPos = 1

		'Beginning of last line
		nLastSepPos = FindToken(vDataSep & tTerm, 1)

		Do Until nDataBoundPos = nLastSepPos
			
			nCurPos = SkipToken(tContentDisp, nDataBoundPos)
			nCurPos = SkipToken(tName, nCurPos)
			sFieldName = ExtractField(tDoubleQuotes, nCurPos)

			nPosFile = FindToken(tFilename, nCurPos)
			nPosBound = FindToken(vDataSep, nCurPos)
			
			If nPosFile <> 0 And  nPosFile < nPosBound Then
				Dim oUploadFile
				Set oUploadFile = New UploadedFile
				
				nCurPos = SkipToken(tFilename, nCurPos)
				auxStr = ExtractField(tDoubleQuotes, nCurPos)
                ' We are interested only in the name of the file, not the whole path
                ' Path separator is \ in windows, / in UNIX
                ' While IE seems to put the whole pathname in the stream, Mozilla seem to 
                ' only put the actual file name, so UNIX paths may be rare. But not impossible.
                osPathSep = "\"
                if InStr(auxStr, osPathSep) = 0 then osPathSep = "/"
				oUploadFile.FileName = Right(auxStr, Len(auxStr)-InStrRev(auxStr, osPathSep))

				if (Len(oUploadFile.FileName) > 0) then 'File field not left empty
					nCurPos = SkipToken(tContentType, nCurPos)
					
                    auxStr = ExtractField(tNewLine, nCurPos)
                    ' NN on UNIX puts things like this in the streaa:
                    '    ?? python py type=?? python application/x-python
					oUploadFile.ContentType = Right(auxStr, Len(auxStr)-InStrRev(auxStr, " "))
					nCurPos = FindToken(tNewLine, nCurPos) + 4 'skip empty line
					
					oUploadFile.Start = nCurPos-1
					oUploadFile.Length = FindToken(vDataSep, nCurPos) - 2 - nCurPos
					
					If oUploadFile.Length > 0 Then UploadedFiles.Add LCase(sFieldName), oUploadFile
				End If
			Else
				Dim nEndOfData
				nCurPos = FindToken(tNewLine, nCurPos) + 4 'skip empty line
				nEndOfData = FindToken(vDataSep, nCurPos) - 2
				If Not FormElements.Exists(LCase(sFieldName)) Then 
					FormElements.Add LCase(sFieldName), String2Byte(MidB(VarArrayBinRequest, nCurPos, nEndOfData-nCurPos))
				else
                    FormElements.Item(LCase(sFieldName))= FormElements.Item(LCase(sFieldName)) & ", " & String2Byte(MidB(VarArrayBinRequest, nCurPos, nEndOfData-nCurPos)) 
                end if 

			End If

			'Advance to next separator
			nDataBoundPos = FindToken(vDataSep, nCurPos)
		Loop
		StreamRequest.Write(VarArrayBinRequest)
	End Sub

	Private Function SkipToken(sToken, nStart)
		SkipToken = InstrB(nStart, VarArrayBinRequest, sToken)
		If SkipToken = 0 then
			Errors.Add "SkipToken = 0", "Error in parsing uploaded binary request."
			Response.End
		end if
		SkipToken = SkipToken + LenB(sToken)
	End Function

	Private Function FindToken(sToken, nStart)
		FindToken = InstrB(nStart, VarArrayBinRequest, sToken)
	End Function

	Private Function ExtractField(sToken, nStart)
		Dim nEnd
		nEnd = InstrB(nStart, VarArrayBinRequest, sToken)
		If nEnd = 0 then
			Errors.Add "nEnd = 0", "Error in parsing uploaded binary request."
			Response.End
		end if
		ExtractField = String2Byte(MidB(VarArrayBinRequest, nStart, nEnd-nStart))
	End Function

	'String to byte string conversion
	Private Function Byte2String(sString)
		Dim i
		For i = 1 to Len(sString)
		   Byte2String = Byte2String & ChrB(AscB(Mid(sString,i,1)))
		Next
	End Function

	'Byte string to string conversion
	Private Function String2Byte(bsString)
		Dim i
		String2Byte =""
		For i = 1 to LenB(bsString)
		   String2Byte = String2Byte & Chr(AscB(MidB(bsString,i,1))) 
		Next
	End Function
End Class

Class UploadedFile
	Public ContentType
	Public Start
	Public Length
	Public Path
	Private nameOfFile

    ' Need to remove characters that are valid in UNIX, but not in Windows
    Public Property Let FileName(fN)
        nameOfFile = fN
        nameOfFile = SubstNoReg(nameOfFile, "\", "_")
        nameOfFile = SubstNoReg(nameOfFile, "/", "_")
        nameOfFile = SubstNoReg(nameOfFile, ":", "_")
        nameOfFile = SubstNoReg(nameOfFile, "*", "_")
        nameOfFile = SubstNoReg(nameOfFile, "?", "_")
        nameOfFile = SubstNoReg(nameOfFile, """", "_")
        nameOfFile = SubstNoReg(nameOfFile, "<", "_")
        nameOfFile = SubstNoReg(nameOfFile, ">", "_")
        nameOfFile = SubstNoReg(nameOfFile, "|", "_")
        nameOfFile = SubstNoReg(nameOfFile, " ", "_")
    End Property

    Public Property Get FileName()
        FileName = nameOfFile
    End Property

    'Public Property Get FileN()ame
End Class


' Does not depend on RegEx, which is not available on older VBScript
' Is not recursive, which means it will not run out of stack space
Function SubstNoReg(initialStr, oldStr, newStr)
    Dim currentPos, oldStrPos, skip
    If IsNull(initialStr) Or Len(initialStr) = 0 Then
        SubstNoReg = ""
    ElseIf IsNull(oldStr) Or Len(oldStr) = 0 Then
        SubstNoReg = initialStr
    Else
        If IsNull(newStr) Then newStr = ""
        currentPos = 1
        oldStrPos = 0
        SubstNoReg = ""
        skip = Len(oldStr)
        Do While currentPos <= Len(initialStr)
            oldStrPos = InStr(currentPos, initialStr, oldStr)
            If oldStrPos = 0 Then
                SubstNoReg = SubstNoReg & Mid(initialStr, currentPos, Len(initialStr) - currentPos + 1)
                currentPos = Len(initialStr) + 1
            Else
                SubstNoReg = SubstNoReg & Mid(initialStr, currentPos, oldStrPos - currentPos) & newStr
                currentPos = oldStrPos + skip
            End If
        Loop
    End If
End Function

sub Resize_AspJpeg(rFilename, rsFilename, rMaxWidth, rMaxHeight, rQuality, rRemoveOrig)

  'newFileName = left(rFilename,instrrev(rFilename,"."))
  ' Create instance of AspJpeg
  Set Jpeg = Server.CreateObject("Persits.Jpeg")
  ' Open source image
  ' Physical path to file
  ' Jpeg.Open "C:\domains\your_domain\images\image.jpg"
  Jpeg.Open rFilename
  'Response.Write(Jpeg.Version)
  jpgver = split(Jpeg.Version,".")(0) & "." & split(Jpeg.Version,".")(1)

  ' New width
  L = 100
  H = 100
  ':: or
  L = rMaxWidth
  H = rMaxHeight
  
  ' Resize, preserve aspect ratio
  'Jpeg.Width = L
  'Jpeg.Height = Jpeg.OriginalHeight * L / Jpeg.OriginalWidth
 if (jpeg.OriginalWidth > rMaxWidth) or (jpeg.OriginalHeight > rMaxHeight) then
  If cint(split(Jpeg.Version,".")(0)) > 0 and cint(split(Jpeg.Version,".")(1)) > 5 then
    jpeg.PreserveAspectRatio = True
	If jpeg.OriginalWidth > jpeg.OriginalHeight Then
   	  jpeg.Width = L
	Else
   	  jpeg.Height = H
	End If

  else
  	If jpeg.OriginalWidth > jpeg.OriginalHeight Then
   	  jpeg.Width = L
   	  jpeg.Height = jpeg.OriginalHeight * L / jpeg.OriginalWidth
	Else
   	  jpeg.Height = H
   	  jpeg.Width = jpeg.OriginalWidth * L / jpeg.OriginalHeight
	End If
  end if
  
  ' create thumbnail and save it to disk
  Jpeg.Save rsFilename
  set Jpeg = nothing
  
  if rRemoveOrig then
    objFSO.deletefile(rFilename)
  end if
 end if
end sub

sub Resize_AspImage(sInFile,sOutFile,maxX,maxY,iQuality,bDeleteOrig)
 Set Image = Server.CreateObject("AspImage.Image")
 if Image.LoadImage(sInFile) then
  Image.GetImageFileSize sInFile, X, Y
  Image.ImageFormat = 1 
  ' 1 = jpg
  ' 2 = bmp
  ' 2 = png
  ' 4 = gif
  Image.JPEGQuality = iQuality 
  
  if X > maxX or Y > maxY then
    if X > maxX then
  	  intYSize = (maxX / X) * Y
  	  Image.ResizeR maxX, intYSize
	else
  	  intXSize = (maxY / Y) * X
  	  Image.ResizeR intXSize, maxY
	end if
  end if
  Image.FileName = sOutFile
  response.Write("<br>" & Image.SaveImage)  
  
  if bDeleteOrig then
    'objFSO.deletefile(sInFile)
  end if
 end if
 Set Image = nothing
end sub

':: get image size and dimensions
function GetBytes(flnm, offset, bytes)
     Dim obFSO
     Dim obFTemp
     Dim obTextStream
     Dim lngSize
     on error resume next
     Set obFSO = CreateObject("Scripting.FileSystemObject")
     ' First, we get the filesize
     Set obFTemp = obFSO.GetFile(flnm)
     lngSize = obFTemp.Size
     set obFTemp = nothing

     fsoForReading = 1
     Set obTextStream = obFSO.OpenTextFile(flnm, fsoForReading)
     if offset > 0 then
        strBuff = obTextStream.Read(offset - 1)
     end if
     if bytes = -1 then		' Get All!
        GetBytes = obTextStream.Read(lngSize)  'ReadAll
     else
        GetBytes = obTextStream.Read(bytes)
     end if
     obTextStream.Close
     set obTextStream = nothing
     set obFSO = nothing
end function

function lngConvert(strTemp)
     lngConvert = clng(asc(left(strTemp, 1)) + ((asc(right(strTemp, 1)) * 256)))
end function

function lngConvert2(strTemp)
     lngConvert2 = clng(asc(right(strTemp, 1)) + ((asc(left(strTemp, 1)) * 256)))
end function

function imgSizeChk(flnm, width, height, depth, strImageType)
     dim strPNG 
     dim strGIF
     dim strBMP
     dim strType
     strType = ""
     strImageType = "(unknown)"
     imgSizeChk = False
     strPNG = chr(137) & chr(80) & chr(78)
     strGIF = "GIF"
     strBMP = chr(66) & chr(77)
     strType = GetBytes(flnm, 0, 3)
     if strType = strGIF then				' is GIF
        strImageType = "GIF"
        Width = lngConvert(GetBytes(flnm, 7, 2))
        Height = lngConvert(GetBytes(flnm, 9, 2))
        Depth = 2 ^ ((asc(GetBytes(flnm, 11, 1)) and 7) + 1)
        imgSizeChk = True
     elseif left(strType, 2) = strBMP then		' is BMP
        strImageType = "BMP"
        Width = lngConvert(GetBytes(flnm, 19, 2))
        Height = lngConvert(GetBytes(flnm, 23, 2))
        Depth = 2 ^ (asc(GetBytes(flnm, 29, 1)))
        imgSizeChk = True
     elseif strType = strPNG then			' Is PNG
        strImageType = "PNG"
        Width = lngConvert2(GetBytes(flnm, 19, 2))
        Height = lngConvert2(GetBytes(flnm, 23, 2))
        Depth = getBytes(flnm, 25, 2)
        select case asc(right(Depth,1))
           case 0
              Depth = 2 ^ (asc(left(Depth, 1)))
              imgSizeChk = True
           case 2
              Depth = 2 ^ (asc(left(Depth, 1)) * 3)
              imgSizeChk = True
           case 3
              Depth = 2 ^ (asc(left(Depth, 1)))  '8
              imgSizeChk = True
           case 4
              Depth = 2 ^ (asc(left(Depth, 1)) * 2)
              imgSizeChk = True
           case 6
              Depth = 2 ^ (asc(left(Depth, 1)) * 4)
              imgSizeChk = True
           case else
              Depth = -1
        end select
     else
        strBuff = GetBytes(flnm, 0, -1)		' Get all bytes from file
        lngSize = len(strBuff)
        flgFound = 0
        strTarget = chr(255) & chr(216) & chr(255)
        flgFound = instr(strBuff, strTarget)
        if flgFound = 0 then
           exit function
        end if
        strImageType = "JPG"
        lngPos = flgFound + 2
        ExitLoop = false
		
        do while ExitLoop = False and lngPos < lngSize
           do while asc(mid(strBuff, lngPos, 1)) = 255 and lngPos < lngSize
              lngPos = lngPos + 1
           loop
           if asc(mid(strBuff, lngPos, 1)) < 192 or asc(mid(strBuff, lngPos, 1)) > 195 then
              lngMarkerSize = lngConvert2(mid(strBuff, lngPos + 1, 2))
              lngPos = lngPos + lngMarkerSize  + 1
           else
              ExitLoop = True
           end if
       loop
       '
       if ExitLoop = False then
          Width = -1
          Height = -1
          Depth = -1
       else
          Height = lngConvert2(mid(strBuff, lngPos + 4, 2))
          Width = lngConvert2(mid(strBuff, lngPos + 6, 2))
          Depth = 2 ^ (asc(mid(strBuff, lngPos + 8, 1)) * 8)
          imgSizeChk = True
       end if
     end if
end function
':::::  END image size functions

%>
<SCRIPT LANGUAGE="VBSCRIPT" RUNAT="SERVER">

Sub ResizeUploadedFiles(ru_path, ru_Suffix, ru_maxWidth, ru_maxHeight, ru_Quality, ru_RemoveOrig, up_filename)
Dim ru_keys, ru_i, ru_curKey, ru_fileName, ru_fso, ru_newFileName, ru_curPath, ru_curName, ru_curExt, ru_lastPos, ru_orgCurPath
  if ru_path <> "" and right(ru_path,1) <> "/" then ru_path = ru_path & "/"
  Set ru_fso = CreateObject("Scripting.FileSystemObject")  
  ru_maxWidth = Cint(ru_maxWidth)
  ru_maxHeight  = Cint(ru_maxHeight)  
          ru_fileName = up_filename
          if ru_fileName <> "" then
            ru_curPath = "" : ru_curName = "" : ru_curExt = ""
            ru_lastPos = InStrRev(ru_fileName,"/")
            if ru_lastPos > 0 then
              ru_curPath = mid(ru_fileName,1,ru_lastPos)	
              ru_curName = mid(ru_fileName,ru_lastPos+1,Len(ru_fileName)-ru_lastPos)	
              ru_fileName = up_filename            
            else
              ru_curName = up_filename	
            end if
            ru_lastPos = InStrRev(ru_curName,".")
            if ru_lastPos > 0 then
              ru_curExt = mid(ru_curName,ru_lastPos+1,Len(ru_curName)-ru_lastPos)	
              ru_curName = mid(ru_curName,1,ru_lastPos-1)
            end if
            ru_curExt = LCase(ru_curExt)
     		ru_orgCurPath = ru_curPath
            if ru_curPath = "" then ru_curPath = ru_path
	    'response.Write("file exist:" & server.MapPath(ru_curPath & up_filename) & "<br>")
            if ru_fso.FileExists(Server.MapPath(ru_curPath & up_filename)) then
                ru_newFileName = ru_curName & ru_Suffix & "." & ru_curExt
	    'response.Write("ru_newFileName:" & ru_newFileName & "<br>")
                FitImage_Comp "image_resizer.aspx", Server.MapPath(ru_CurPath & ru_fileName), Server.MapPath(ru_curPath & ru_newFileName), ru_maxWidth, ru_maxHeight, ru_Quality
                if ru_RemoveOrig then
                  if LCase(ru_fileName) <> LCase(ru_newFileName) then
                    ru_fso.DeleteFile Server.MapPath(ru_curPath & ru_fileName)
                  end if  
                end if
            end if
          end if	
End Sub

sub FitImage_Comp(DotNetResize,imgFile,newImgFile,maxWidth,maxHeight,Quality)
    select case DetectDotNetComponent(DotNetResize)
    case "DOTNET1"
      Image_Size_DotNet "Msxml2.ServerXMLHTTP.4.0",DotNetResize,imgFile,newImgFile,maxWidth,maxHeight,Quality
    case "DOTNET2"
      Image_Size_DotNet "Msxml2.ServerXMLHTTP",DotNetResize,imgFile,newImgFile,maxWidth,maxHeight,Quality
    case "DOTNET3"
      Image_Size_DotNet "Microsoft.XMLHTTP",DotNetResize,imgFile,newImgFile,maxWidth,maxHeight,Quality
    end select
end sub

function DetectDotNetComponent(DotNetResize)
  Dim DotNetImageComponent, ResizeComUrl, LastPath
	
	DotNetImageComponent = ""
	ResizeComUrl = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("PATH_INFO")
	ResizeComUrl = replace(ResizeComUrl,"tiny_mce/plugins/advimage/","")
	LastPath = InStrRev(ResizeComUrl,"/")
	if LastPath > 0 then
		ResizeComUrl = left(ResizeComUrl,Lastpath)
	end if
	ResizeComUrl = ResizeComUrl & DotNetResize
	'Response.Write ResizeComUrl & "<br>"
	
	'Check for ASP.NET
	if DotNetCheckComponent("Microsoft.XMLHTTP", ResizeComUrl) = true then
	  'Response.Write "FOUND: ASP.NET Microsoft.XMLHTTP<br>"
	  DotNetImageComponent = "DOTNET3"
	else
	  if DotNetCheckComponent("Msxml2.ServerXMLHTTP", ResizeComUrl) = true then
		DotNetImageComponent = "DOTNET2"
	  else
		if DotNetCheckComponent("Msxml2.ServerXMLHTTP.4.0", ResizeComUrl) = true then
		  DotNetImageComponent = "DOTNET1"
		else
		  Response.Write "NOT FOUND: ASP.NET Server Component<br>"
		end if
	  end if
	end if
	'on error goto 0
  
	DetectDotNetComponent = DotNetImageComponent
end function

function DotNetCheckComponent(DotNetObj, ResizeComUrl)
  dim objHttp, Detection
	Detection = false
  on error resume next
  err.clear
  Set objHttp = Server.CreateObject(DotNetObj)
  if err.number = 0 then
    objHttp.open "GET", ResizeComUrl, false
	if err.number = 0 then
      objHttp.Send ""
	  if (objHttp.status <> 200 ) then
		Response.Write "An error has accured with ASP.NET component " & DotNetObj & "<br>;"
		Response.Write "Error:<br>" & objHttp.responseText & "<br>"
		Response.End
	  end if
      if trim(objHttp.responseText) <> "" and trim(objHttp.responseText) = "DONE" then
        Detection = true
      end if
	end if
  End if
  Set objHttp = nothing
  on error goto 0
  DotNetCheckComponent = Detection
end function

sub Image_Size_DotNet(DotNetComp, DotNetResize, imgFile,newImgFile,maxWidth,maxHeight,Quality)
  Dim objHttp, ResizeComUrl, ResizeParams, LastPath
  'Response.Write "Image_Size_DotNet<br>"
  ResizeParams = "?f=" & Server.UrlEncode(imgFile) & "&nf=" & Server.UrlEncode(newImgFile) & "&w=" & maxWidth & "&h=" & maxHeight & "&q=" & Quality
  ResizeComUrl = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("PATH_INFO")
	ResizeComUrl = replace(ResizeComUrl,"tiny_mce/plugins/advimage/","")
  LastPath = InStrRev(ResizeComUrl,"/")
  if LastPath > 0 then
    ResizeComUrl = left(ResizeComUrl,Lastpath)
  end if
  ResizeComUrl = ResizeComUrl & DotNetResize & ResizeParams
  'Response.Write ResizeComUrl & "<br>"

  on error resume next
  set objHttp = Server.CreateObject(DotNetComp)
  if err.number <> 0 then
    Response.Write "ERROR: ASP.NET (" & DotNetComp & ") is not installed!<br>Image resize is not available"
    Response.End
  end if
  
  objHttp.open "GET", ResizeComUrl, false
  objHttp.Send ""
  
  ' Check notification validation
  if (objHttp.status <> 200 ) then
    ' HTTP error handling
    Response.Write "HTTP ERROR: " & objHttp.status & "<br>"
    Response.Write "Returned:<br>" & objHttp.responseText 
    
  elseif (objHttp.responseText = "Done") then
  'Response.Write "it says DONE<br>"
  else
    if trim(objHttp.responseText)="" or instr(objHttp.responseText,"@ Page Language=""C#""")>0 then
      Response.Write "DOT NET Unsupported"
	else
  	  'Response.Write "unspecified error: " & objHttp.responseText & "<br>"
    end if
  end if
  Set objHttp = Nothing
  on error goto 0
end sub

</SCRIPT>