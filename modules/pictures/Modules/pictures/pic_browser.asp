<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'<> Copyright (C) 2005-2008 Dogg Software All Rights Reserved
'<>
'<> By using this program, you are agreeing to the terms of the
'<> SkyPortal End-User License Agreement.
'<>
'<> All copyright notices regarding SkyPortal must remain 
'<> intact in the scripts and in the outputted HTML.
'<> The "powered by" text/logo with a link back to 
'<> http://www.SkyPortal.net in the footer of the pages MUST
'<> remain visible when the pages are viewed on the internet or intranet.
'<>
'<> Support can be obtained from support forums at:
'<> http://www.SkyPortal.net
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

response.buffer = true
on error resume next
%>
<!-- #include file="pic_img.asp" -->
<!-- #include file="pic_config.asp" -->
<%
Dim objFSO, objFolder, objItem


varDirPath = Request.QueryString("dirpath")
if varDirPath = "" then varDirPath = strUpDirPath
if left(varDirPath,len(BasePath)) <> BasePath then varDirPath = strUpDirPath

If Len(varDirPath) > 1 Then
   strPath = varDirPath & "/"
Else
   strPath = BasePath & "/"
   'BasePath = strPath 
End If

function pgCss() %>
<style type="text/css">
#previewField {
	margin:0px;
	border: 1px solid black;
	width: 99%;
	height:250px;
	overflow: auto;
}

#files_list {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 12px;
	padding: 3px;
	border: 1px solid black;
}

.leftColumn
{float:left;width: 60%;}

.rightColumn
{float: right;width: 38%;}

.rightColumn img
{border: 0px none;}

.panel_wrapper div.current {
	height: 310px;
/*	border: 1px solid red; */
}

.explorer
{
	height: 250px;;
	width: 98%;;
	border: 1px solid black;
	background-color: white;
	overflow : scroll;
}

.explorer table
{
	width:99%;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 11px;
}
.explorer td
{padding: 0px 5px;}

.darkRow
{background-color: #F0F0EE;}

.explorer th
{font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 11px;}

.explorer table a:link {color: #000; text-decoration:none;}
.explorer table a:active {color: #000; text-decoration:none;}
.explorer table a:visited {color: #000; text-decoration:none;}
.explorer table a:hover {color: blue; text-decoration:none;}
</style>
  <%
end function

Function ExtractDirName(strFilename)
    ' Removes the directory from a string that contains path and filename
    If Len(strFilename) > 1 Then
		For I = 1 to 2
			Dim X 
			For X = Len(strFilename) To 1  Step -1 
				If Mid(strFilename, X, 1) = "/" Then Exit  For 
			Next 
			If Len(strFilename) > 1 Then
				strFilename = Left(strFilename, X - 1)
			Else
				strFilename = "/"
				X = 2
			End If
		Next 
	   	ExtractDirName = Left(strFilename, X - 1)
    End If
End Function

Function ShowImageForType(strName)
' For the 'explorer': shows file-icons
	Dim strTemp
	strTemp = strName
	If strTemp <> "dir" Then
		strTemp = LCase(Right(strTemp, Len(strTemp) - InStrRev(strTemp, ".", -1, 1)))
	End If
	Select Case strTemp
		Case "asp"
			strTemp = "asp"
		Case "dir"
			strTemp = "dir"
		Case "htm", "html"
			strTemp = "htm"
		Case "gif", "jpg"
			strTemp = "img"
		Case "txt"
			strTemp = "txt"
		Case Else
			strTemp = "misc"
	End Select
	strTemp = "<img src=""" & strImgDir & "dir_" & strTemp & ".gif"" width=""16"" height=""16"" border=""0"" alt=""" & strTemp & """ />"
	ShowImageForType = strTemp
End Function

function SaveFiles(PathToSaveTo)
' Saves potentially uploaded files
    Dim Upload, fileName, fileSize, ks, i, fileKey
    Set Upload = New FreeASPUpload
    Upload.Save(PathToSaveTo)
    SaveFiles = ""
	
    ks = Upload.Errors.keys
	' if errors ar returned by the component
    if (UBound(ks) <> -1) then
        SaveFiles = ""
        for each fileKey in Upload.Errors.keys
            SaveFiles = SaveFiles & Upload.Errors(fileKey)&"\n" 
        next
    else
      kf = Upload.UploadedFiles.keys
      if (UBound(kf) <> -1) then
        'SaveFiles = "Upload response:\n\n"
        for each fileKey in Upload.UploadedFiles.keys
		  if instr(Upload.UploadedFiles(fileKey).ContentType,"image") <> 0 then
		    if Upload.UploadedFiles(fileKey).Length > iMaxSize then
			  ':: FILE SIZE TO LARGE
		      if Right(PathToSaveTo, 1) <> "\" then PathToSaveTo = PathToSaveTo & "\"
		      SaveFiles = SaveFiles & "File too large:\n"
		      SaveFiles = SaveFiles & "/nMax size: " & iMaxSize & " bytes\n"
		      SaveFiles = SaveFiles & "/nYour file: " & Upload.UploadedFiles(fileKey).Length & " bytes\n"
		      'SaveFiles = SaveFiles & "\n" & replace(PathToSaveTo,"\","\\") & Upload.UploadedFiles(fileKey).FileName
			  Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
			  oFSO.deletefile(PathToSaveTo & Upload.UploadedFiles(fileKey).FileName)
			  set oFSO = nothing
			else
			  if Upload.UploadedFiles(fileKey).Length <= 1024 then
                SaveFiles = SaveFiles & "Success: " & Upload.UploadedFiles(fileKey).FileName & " \n(" & Upload.UploadedFiles(fileKey).Length & " bytes)\n"
			  else
                SaveFiles = SaveFiles & "Success: " & Upload.UploadedFiles(fileKey).FileName & " \n(" & clng(Upload.UploadedFiles(fileKey).Length/1024) & " kb)\n"
			  end if
			
			  ':: do resizing here
			  if iMaxWidth > 0 and iMaxHeight > 0 then
   			    select case lcase(strImgComp)
    	  	      case "aspnet"
				    fname = Upload.UploadedFiles(fileKey).FileName
				    'Response.Write("Start Resize<br>")
		    	    ResizeUploadedFiles varDirPath,"_rs",iMaxWidth,iMaxHeight,90,true,fname
				    'Response.Write("Finish Resize<br>")
    	  	      case "aspjpeg"
		    	    if Right(PathToSaveTo, 1) <> "\" then PathToSaveTo = PathToSaveTo & "\"
   'Resize_AspJpeg(rFilename, rsFilename, rMaxWidth, rMaxHeight, rQuality, rRemoveOrig)
   				    orig = PathToSaveTo & Upload.UploadedFiles(fileKey).FileName
				    newName = PathToSaveTo & "r_" & replace(Upload.UploadedFiles(fileKey).FileName," ","_")
		      	    Resize_AspJpeg orig,newName,iMaxWidth,iMaxHeight,90,true
    	  	      case "aspimage"
		    	    if Right(PathToSaveTo, 1) <> "\" then PathToSaveTo = PathToSaveTo & "\"
   				    orig = PathToSaveTo & Upload.UploadedFiles(fileKey).FileName
				    newName = PathToSaveTo & "r_" & replace(Upload.UploadedFiles(fileKey).FileName," ","_")
		      	    Resize_AspImage orig,newName,iMaxWidth,iMaxHeight,90,true
  			    end select
			  end if
			end if
		  else
		    if Right(PathToSaveTo, 1) <> "\" then PathToSaveTo = PathToSaveTo & "\"
		    SaveFiles = SaveFiles & "Illegal file type:\n" & Upload.UploadedFiles(fileKey).FileName
		    'SaveFiles = SaveFiles & "\n" & replace(PathToSaveTo,"\","\\") & Upload.UploadedFiles(fileKey).FileName
			Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
			oFSO.deletefile(PathToSaveTo & Upload.UploadedFiles(fileKey).FileName)
			set oFSO = nothing
		  end if
        next
      else
        SaveFiles = "The file name specified in the upload form does not correspond to a valid file in the system."
      end if
	
        'SaveFiles = "no errors"
    end if
end function
%>

<%' Now to the Runtime code:
'Response.Write(Server.MapPath(BasePath))
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if not objFSO.folderexists(Server.MapPath(strUpDirPath)) then
  objFSO.createfolder(Server.MapPath(strUpDirPath))
end if
Set objFolder = objFSO.GetFolder(Server.MapPath(strPath))
%>
<html>
<head>
	<title>Mass Pictures to DB</title>
	<script language="javascript" type="text/javascript" src="multiSelect.js"></script>
<script type="text/javascript">
var maxImgWidth = 260;
var maxImgHeight = 250;

function PreviewFile()
{
 // resizes the preview-image to fit the window
  var Img = new Image();
  Img.src = document.images['ImgVar'].src
  var x=parseInt(Img.width);
  var y=parseInt(Img.height);
  if (x>maxImgWidth) {
    y*=maxImgWidth/x;
    x=maxImgWidth;
  }
  if (y>maxImgHeight) {
    x*=maxImgHeight/y;
    y=maxImgHeight;
  }
 if ((x!=0)||(y!=0))
	 {
	  document.images['ImgVar'].width=x;
	  document.images['ImgVar'].height=y;
	 }
}

function previewImg(src){
	//alert(src);
	var eTxt;
	eTxt = '<a href="' + src + '" target="_blank">';
	eTxt += '<img id="ImgVar" src="' + src + '" border="0" alt="" title="Click for full size image" onload="PreviewFile()" />';
	eTxt += '</a>';
	var elm = document.getElementById('previewField');
		elm.innerHTML = eTxt;
}

function previewUFile(src){
	//alert(src);
	var eTxt;
	eTxt = '<a href="' + src + '" target="_blank">';
	eTxt += '<img id="ImgVar" src="' + src + '" border="0" alt="" title="Click for full size image" onload="PreviewFile()" />';
	eTxt += '</a>';
	var elm = document.getElementById('previewField');
		elm.innerHTML = eTxt;
}

</script>
  <% pgCss() %>
 <base target="_self" />
</head>
<body>
<%
' This piece of code displays any errors from the upload component in a JS popup.

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	SaveFiles(Server.MapPath(varDirPath))
	If varDirPath <> "" Then
		'response.write("alert(""Upload Result:\n" & SaveFiles(Server.MapPath(varDirPath)) & """);")
	Else
		'varDirPath = "/"
		'response.write("alert(""Upload Result 2\n: " & SaveFiles(Server.MapPath(varDirPath)) & """);")
	End If
end if
%>
<form name="filebrowser" method="POST" enctype="multipart/form-data" action="" onsubmit="return finished();">
<div class="panel_wrapper">
	<div id="general_panel" class="panel current">
		<div class="leftColumn">
			<h3>Image browser</h3>
		<%strNPath = strPath%>
		<% Response.flush() %>
			<div class="explorer">
					<table>
						<tr>
							<th align="left" colspan="3">
							<%= ShowImageForType("dir") %>&nbsp;
							<%= strNPath %></th>
							<!--<th>Modified</th>-->
						</tr>
						<tr>
							<th>Name</th>
							<th>Size</th>
							<th>Type</th>
							<!--<th>Modified</th>-->
						</tr>
						<!--<tr>
							<td>&nbsp;<a href="?dirpath=<%'=ExtractDirName(strNPath)%>"><img src="images/dir_pdir.gif" width="16" height="16" border="0" alt="parent directory">..</a></td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
						</tr>-->
		<%			Dim RowCount
						RowCount = 0
					For Each objItem In objFolder.SubFolders
						If InStr(1, objItem, "_vti", 1) = 0 Then%>
						<tr <% If RowCount MOD 2 = 0 Then%>class="darkRow"<% End If %>>
							<td><%= ShowImageForType("dir") %>&nbsp;<a href="?dirpath=<%= strPath & objItem.Name %>"><%= objItem.Name %></a></td>
							<td><%'= objItem.Size %></td>
							<td><%'= objItem.Type %></td>
							<!--<td><%'= objItem.DateCreated %></td> -->
						</tr>
		<%				RowCount = RowCount + 1
						End If
					Next %>
		
		<%			For Each objItem In objFolder.Files	%>
						<tr <% If RowCount MOD 2 = 0 Then%>class="darkRow"<% End If %>>
							<td nowrap>
							<input type="hidden" name="imgVar<%= RowCount %>" id="imgVar<%= RowCount %>" value="<%= strImgPath & objItem.Name %>" /><%= ShowImageForType(objItem.Name) %>&nbsp;<a href="Javascript:;" onClick="previewImg('<%= BasePath & "/" & objItem.Name %>')"><%= objItem.Name %></a></td>
							<td><%= objItem.Size %></td>
							<td><%= objItem.Type %></td>
							<!--<td nowrap><%'= 'objItem.DateCreated %></td>
							<a onclick="preview(this,'100','100')" href="Javascript:FileChosen('<%'= objItem.Name %>')">-->
						</tr>
		<%				RowCount = RowCount + 1
					Next %>
					</table>
			</div>
		</div>
		<%
		Set objItem = Nothing
		Set objFolder = Nothing
		Set objFSO = Nothing
		%>
		<div class="rightColumn">
			<h3>Image Preview</h3>
			<div id="previewField">
			<img src="/images/spacer.gif" width="1" height="1">
			</div>
		</div>
		
		<br clear="all"/><br/>
<!-- This is where the output will appear -->
<div id="files_list">
		<table width="100%" border="0" cellspacing="4" cellpadding="0">
			<tr>
				<td colspan="2">
				<b><font size="3" face="Arial, Helvetica, sans-serif">
				<label for="file_element">Select files to be uploaded:
				</label></font></b>
				</td>
			</tr>
			<tr>
				<td width="274">
				<input type="hidden" name="chosendir" value="<%= strUpDirPath %>" />
				<input type="file" name="file_element" id="file_element" size="32" /></td>
				<td>
				<input id="insert" type="submit" name="Submit" value="Upload" /></td>
			</tr>
		</table>

</div>
<script type="text/javascript">
	<!-- Create an instance of the multiSelector class, pass it the output target and the max number of files -->
	var multi_selector = new MultiSelector( document.getElementById( 'files_list' ), 0 );
	<!-- Pass in the file element -->
	multi_selector.addElement( document.getElementById( 'file_element' ) );
</script>
  </div>
</div>
</form>
</body> 
</html> 
<% 
on error goto 0 %>
