<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: doctree.asp
' AUTHOR: Steve Loar	
' CREATED: 3/26/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Documents.
'
' MODIFICATION HISTORY
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iFolderCount, aPath, bHavePath

If request("path") <> "" Then
	If InStr(UCase(request("path")),"DECLARE") > 0 And InStr(UCase(request("path")),"VARCHAR") > 0 And InStr(UCase(request("path")),"EXEC") > 0 Then 
		response.redirect "doctree.asp"
	End If 
	If InStr(UCase(request("path")),"SELECT") > 0 And InStr(UCase(request("path")),"VARCHAR") > 0 And InStr(UCase(request("path")),"CAST") > 0 Then 
		response.redirect "doctree.asp"
	End If
	bHavePath = True 
	aPath = Split(UCase(request("path")),"/")
	iPathSize = Ubound(aPath)
Else
	bHavePath = False 
End If 
 

Session("RedirectPage") = Request.ServerVariables("SCRIPT_NAME") 
Session("RedirectLang") = "Return to Documents"


%>

<html>
<head>
	<title>E-Gov Services - <%=sOrgName%></title>

	<!-- Required CSS -->
	<link rel="stylesheet" type="text/css" href="../yui/build/tabview/assets/skins/sam/tabview.css" />
	<link rel="stylesheet" type="text/css" href="../yui/build/treeview/assets/skins/sam/treeview.css">
	<link rel="stylesheet" type="text/css" href="../yui/examples/treeview/assets/css/folders/tree.css"> 
	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

	<!-- Dependency source file -->
	<script type="text/javascript" src="../yui/build/yahoo-dom-event/yahoo-dom-event.js" ></script>
	<script type="text/javascript" src="../yui/build/animation/animation-min.js"></script>
	<script type="text/javascript" src="../yui/build/element/element-beta.js"></script>

	<!-- TreeView source file -->
	<script src = "../yui/build/treeview/treeview-min.js" ></script>  
	
	<script type="text/javascript">
	<!--
		var tree;
	//-->
	</script>

	<style>
		#expandcontractdiv {border:1px dotted #dedede; background-color:#EBE4F2; margin:0 0 .5em 0; padding:0.4em;}
		#treeDiv1 { background: #fff; padding:1em; margin-top:1em; }
	</style>

	<!--#Include file="../include_top.asp"-->

	<tr><td valign="top">
		<p>
			<font class="pagetitle"><%=sOrgName%> Online Documents</font><br />
		</p>

	<!--BEGIN:  USER REGISTRATION - USER MENU-->
	<%	RegisteredUserDisplay( "../" ) %>
	<!--END:  USER REGISTRATION - USER MENU-->


	<!--BEGIN: ADOBE CODE-->
	<div style="width:750px; padding-left:10px;padding-bottom:10px;">Some of the pages within this section link to Portable Document Format (PDF) files which require a PDF reader to view. You may download a free copy of Adobe&reg; Reader&reg; if you do not already have it on your computer.<br /><br />
		<a href='http://www.adobe.com/products/acrobat/readstep2.html' target='_blank' title='Get Adobe Acrobat Reader Plug-in Here'><img border=0 src="../images/adreader.gif" hspace=10>Get Adobe Reader.</a>
	</div>
	<!--END: ADOBE CODE-->

	<div>
	<a href="javascript:tree.expandAll()">Expand all</a>
    <a href="javascript:tree.collapseAll()">Collapse all</a>
	</div>

  <table border=0>
  <tr>
  <td valign="top" style="padding-left:10px;">
   <form action="../search/search.asp" method="post" id="form1" name="frmSearch">
	<p><font class="searchlabel">Search Documents:</font><br />
		<input type="hidden" name="Action" value="Go" />
		<input type="text" id="SearchString" name="SearchString" size="65" maxlength="100" style="background-color:#eeeeee; border:1px solid #000000; width:144px;" /><br />
		<div class="quicklink" align="right">
			<a href="#" onClick='ValidateSearch();'><img src="menu/images/go.gif" border="0" /><font class="searchlink">Search</font></a>&nbsp;&nbsp;
		</div>
	        
	</p>
   </form>
  </td>
  <td valign="top">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </td>
  <td valign="top">
	<!--BEGIN: DISPLAY DOCUMENT/FOLDER TREE-->
	<%

'	Dim i
'	If bHavePath Then 
'		For i = 0 to Ubound(aPath)
'			If i > 3 Then 
'				Response.Write i & " = " & aPath(i) & "<br />"
'			End If 
'		Next
'		response.write iPathSize & "<br />"
'	End If 

	' GENERATE DOCUMENT/FOLDER TREE

	' GET CITY NAME
	sLocationName =  GetVirtualDirectyName()
	'response.write sLocationName & "<br />"

	'Set FSO = CreateObject("Scripting.FileSystemObject")

	' BEGIN FOLDER TREE
	
	response.write vbcrlf & "<div id=""treeDiv1"">"
	response.write vbcrlf & "<ul>"

	' LIST ALL FOLDERS AND DOCUMENTS
	'ENUMERATEDOCUMENTS FSO.GetFolder(Server.Mappath("/public_documents300/" & sLocationName & "/published_documents/")),Server.Mappath("/public_documents300/" & sLocationName & "/published_documents/"),"public_documents300/" & sLocationName & "/published_documents"  
	iFolderId = GetFolderId( iOrgId, "published_documents" )
	'response.write "iFolderId = " & iFolderId & "<br /><br />"

	Dim iNodeCount, aFolders
	iNodeCount = 0
	ReDim aFolders(3)

	iFolderCount = ShowFoldersAndDocs( iOrgId, iFolderId, iNodeCount, True, 3 )

	response.write vbcrlf & "</ul>"
	response.write vbcrlf & "</div>"

'	If bHavePath Then 
'		response.write " Ubound(aFolders) " & Ubound(aFolders) & "<br />"
'		For i = 0 to Ubound(aFolders)
'			Response.Write i & " = " & aFolders(i) & "<br />"
'		Next
'	End If 
	%>

	<% ' DESTROY OBJECTS
	   'SET FSO = NOTHING
	%>

	<!--END: DISPLAY DOCUMENT/FOLDER TREE-->
  </td>
  </tr>
  </table>

<script>
<!--
	treeInit();


	function treeInit() 
	{ 
		tree = new YAHOO.widget.TreeView("treeDiv1");   
		tree.render();  

		var nodeToExpand;
<%		
'		If Ubound(aFolders) > 3 Then 
'			For i = 4 To Ubound(aFolders)
'				response.write vbcrlf & "nodeToExpand = tree.getNodeByIndex( " & aFolders(i) & " );"
'				response.write vbcrlf & "nodeToExpand.expand();"
'			Next 
'		End If 
%>
	}   
//-->
</script>



<!--#Include file="../include_bottom.asp"-->


<%


'-------------------------------------------------------------------------------------------------------
' Sub ShowFoldersAndDocs( iOrgId, iFolderId, iNodeCount )
'-------------------------------------------------------------------------------------------------------
Function ShowFoldersAndDocs( ByVal iOrgId, ByVal iFolderId, ByRef iNodeCount, ByVal bParentMatch, ByVal iFolderMatchLevel )
	Dim sSql, oRs, iFolderCount, iSubFolderCount, iMatchLevel, bMatch

	iFolderCount = 0 
	iSubFolderCount = 0
	bMatch = False 
	iMatchLevel = 0

	' Get the set of subfolders, loop and display
	sSql = "SELECT folderid, foldername FROM DocumentFolders WHERE orgid = " & iOrgId & " AND ParentFolderID = " & iFolderId & " ORDER BY foldername"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		If UserHasDocFolderAccess( iOrgId, request.Cookies("userid"), oRs("folderid") ) Then
			iFolderCount = iFolderCount + 1
			iNodeCount = iNodeCount + 1
			response.write vbcrlf & "<li id=""foldheader""><strong> &nbsp;" & UCase(oRs("foldername")) '& "-" & iNodeCount & 
			response.write "</strong></li>"
'			If iNodeCount = 62 Then
'				response.End 
'			End If
			response.write vbcrlf & "<ul id=""foldinglist"" name=""foldinglist"" style=""display:none;"" >"
			If bHavePath Then 
				iMatchLevel = iFolderMatchLevel + 1
				If iMatchLevel <= iPathSize And iFolderMatchLevel > 0 Then 
					If UCase(aPath(iMatchLevel)) = UCase(Trim(oRs("foldername"))) Then
						If bParentMatch Then 
							bMatch = True 
							Redim Preserve aFolders(iMatchLevel)
							aFolders(iMatchLevel) = iNodeCount
						End If 
					End If 
				End If 
			End If 
			iSubFolderCount = ShowFoldersAndDocs( iOrgId, oRs("folderid"), iNodeCount, bMatch, iMatchLevel )
			ShowDocumentsInFolder oRs("folderid"), iSubFolderCount, iNodeCount
			response.write vbcrlf & vbcrlf & "</ul>"
		End If 
		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 
	ShowFoldersAndDocs = iFolderCount
End Function  


'-------------------------------------------------------------------------------------------------------
' Sub ShowDocumentsInFolder( iFolderId, iFolderCount, iNodeCount )
'-------------------------------------------------------------------------------------------------------
Sub ShowDocumentsInFolder( iFolderId, iFolderCount, ByRef iNodeCount )
	Dim sSql, oRs, sDocumentSize

	' Get the set of files, loop and display
	sSql = "SELECT DocumentTitle, DocumentURL, documentsize FROM Documents WHERE ParentFolderID = " & iFolderId & " ORDER BY DocumentTitle"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			sFoldersPath = Replace( Request.ServerVariables("SERVER_NAME") & oRs("DocumentURL"), "/custom/pub", "" )
			If CLng(oRs("documentsize")) > CLng(1024) Then
				sDocumentSize = FormatNumber((oRs("documentsize") / 1024),0)  & " KB"
			Else
				sDocumentSize = FormatNumber(oRs("documentsize"),0) & " Bytes"
			End If
			response.write vbcrlf & "<li>"
			response.write "<a class=""documentlist"" TARGET=""DOCUMENTS"" href=""" & "http://" & sFoldersPath  & """ > "
			'response.write vbcrlf & "<img src=""" & GetFileImageURL( Trim(oRs("DocumentTitle")) ) & """ border=""0"" /> " 
			iNodeCount = iNodeCount + 1
			response.write vbcrlf & UCase(Trim(oRs("DocumentTitle"))) & " ( " & sDocumentSize & " ) " '& "-" & iNodeCount
'			If iNodeCount = 62 Then
'				response.End 
'			End If 
			response.write vbcrlf & "</a></li>"
			oRs.MoveNext
		Loop
	Else
		iNodeCount = iNodeCount + 1
		If CLng(iFolderCount ) = CLng(0) Then 
			response.write vbcrlf & "<li class=""emptyfolder"" style=""font-style: italic;"">(EMPTY)" & "-" & iNodeCount & "</li>"
		End If 
	End If 
End Sub 


'-------------------------------------------------------------------------------------------------------
' SUB ENUMERATEDOCUMENTS(FOLDER)
'-------------------------------------------------------------------------------------------------------
Sub ENUMERATEDOCUMENTS(Folder,sPath,sVpath)

	' BUILD HYPERLINK BASE PATH
	sVirtualpath = replace(Folder.Path,sPath,"")
	sTempPath =  replace(Folder.Path,sPath,"")
	sVirtualpath = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & sVpath & replace(sVirtualpath,"\","/")

	' LIST CONTENTS OF FOLDER (SUBFOLDERS AND FILES)
	For Each SubFolder in Folder.SubFolders

		sVirtualpath2 =  replace(sVpath,"public_documents300","/public_documents300/custom/pub") & replace(sTempPath,"\","/") & "/" & SubFolder.Name

		' CHECK SECURITY FOR ACCESS
		If HasAccess(iorgid,request.Cookies("userid"),sVirtualpath2) Then

			' WRITE FOLDER INFORMATION
			response.write vbcrlf
			response.write "<li id=""foldheader""> " & SubFolder.Name & "</li>" & vbcrlf
			response.write "<ul id=""foldinglist"" name=""foldinglist"" style=""display:none;"" >" & vbcrlf

			' RECURSIVE CALL TO GET ANY SUBFOLDERS OF THE CURRENT FOLDER
			ENUMERATEDOCUMENTS Subfolder, sPath, sVpath

			' LIST FILES IN THE CURRENT FOLDER
			For each File in Subfolder.Files
				' GET FILE SIZE
				If File.Size > 1024 Then
					sFileSize = FormatNumber((File.Size / 1024),0)  & " KB"
				Else
					sFileSize =  FormatNumber(File.Size,0) & " Bytes"
				End If

				sHyperlink = sVirtualPath & "/" & Subfolder.Name & "/" & File.Name
				'response.write  "<li class=""" & GetListItemClassName( File.Name ) & """><a class=""documentlist"" TARGET=""DOCUMENTS"" href=""" & sHyperlink  & """ > " & UCASE(File.Name) & " ( " & sFileSize & ") </a></li>"
				response.write vbcrlf & "<li><a class=""documentlist"" TARGET=""DOCUMENTS"" href=""" & sHyperlink  & """ > " & UCASE(File.Name) & " ( " & sFileSize & ") </a></li>"
			Next

			' IF FOLDER CONTAINS NO FILES OR FOLDERS SHOW AS EMPTY
			If (Subfolder.Files.Count < 1) AND (Subfolder.SubFolders.Count < 1) Then
				'SET MARGIN TO -20 TO ALIGN SUB FOLDERS AND DOCUMENTS WITHIN A FOLDER PROPERLY.
				response.write  "<li class=""emptyfolder"" STYLE=""margin-left:-20px;""> (<i> EMPTY </i>) </li>"  & vbcrlf
			End IF

			' CLOSE FOLDER LIST TAG
			response.write "</ul>"
		Else 
			' Added to allow direct folder links to work while skipping restricted access folders
			response.write "<li id=""foldheader"" style=""display:none;"" ></li>" & vbcrlf
			response.write "<ul id=""foldinglist"" name=""foldinglist"" style=""display:none;"" ></ul>" & vbcrlf
			
			' RECURSIVE CALL TO GET ANY SUBFOLDERS OF THE CURRENT FOLDER
			ENUMERATEDOCUMENTS Subfolder, sPath, sVpath
		End If
	Next

End Sub


'-------------------------------------------------------------------------------------------------------
' Function UserHasDocFolderAccess( iOrgId, iUserId, iFolderId )
'-------------------------------------------------------------------------------------------------------
Function UserHasDocFolderAccess( iOrgId, iUserId, iFolderId )
	Dim rstAccess, sSql, iReturnValue

	On Error Resume Next

	iReturnValue = False

	Set oCnn = Server.CreateObject("ADODB.Connection")
	oCnn.Open Application("DSN")
	sSql = "EXEC CheckFolderIdAccess '" & iOrgId & "','" & iUserId & "','" & iFolderId & "'"
	'response.write sSql & "<br /><br />"
	Set rstAccess = oCnn.Execute(sSql)

	If Not rstAccess.EOF Then
		If rstAccess("folderid") = iFolderId Then
			iReturnValue = True
		End If
	End If

	oCnn.Close
	Set rstAccess = Nothing
	Set oCnn = Nothing

	UserHasDocFolderAccess = iReturnValue

End Function


'-------------------------------------------------------------------------------------------------------
' Function GetListItemClassName( sName )
'-------------------------------------------------------------------------------------------------------
Function GetListItemClassName( sName )

	sReturnValue = "msie"

	Select Case LCase(Right(Trim(sName),3))
		Case "doc"
            sReturnValue = "doc"
        Case "xls"
            sReturnValue = "xls"
        Case "ppt"
            sReturnValue = "ppt"
        Case "htm"
            sReturnValue = "htm"
        Case "pdf"
            sReturnValue = "pdf"
		Case "gif"
            sReturnValue = "gif"
        Case "jpg"
            sReturnValue = "jpg"
	End Select

	  GetListItemClassName = sReturnValue

End Function


'-------------------------------------------------------------------------------------------------------
' Function GetFileImageURL( sFileName )
'-------------------------------------------------------------------------------------------------------
Function GetFileImageURL( sFileName )

	sReturnValue = "./menu/images/"

	Select Case LCase(Right(Trim(sFileName),3))
		Case "doc"
            sReturnValue = sReturnValue & "msword.gif"
        Case "xls"
            sReturnValue = sReturnValue & "msexcel.gif"
        Case "ppt"
            sReturnValue = sReturnValue & "msppt.gif"
        Case "htm"
            sReturnValue = sReturnValue & "msie.gif"
        Case "pdf"
            sReturnValue = sReturnValue & "pdf.gif"
		Case "gif"
            sReturnValue = sReturnValue & "imageicon.gif"
        Case "jpg"
            sReturnValue = sReturnValue & "imageicon.gif"
		Case Else 
			sReturnValue = sReturnValue & "document.gif"
	End Select

	  GetFileImageURL = sReturnValue

End Function


'-------------------------------------------------------------------------------------------------------
' FUNCTION HASACCESS(IORGID,IUSERID,STRVPATH)
'-------------------------------------------------------------------------------------------------------
Function HasAccess(iorgid,iuserid,strvpath)

	  On Error Resume Next

	  iReturnValue = False

	  Set oCnn = Server.CreateObject("ADODB.Connection")
	  oCnn.Open Application("DSN")
	  sSql = "EXEC CHECKFOLDERACCESS '" & iorgid & "','" & iuserid & "','" & strvpath & "'"
	  Set rstAccess = oCnn.Execute(sSql)

	  If NOT rstAccess.EOF Then
		If rstAccess("folderid") >= 0 Then
			iReturnValue = True
		End If
	  End If

	  oCnn.Close
	  Set rstAccess = Nothing
	  Set oCnn = Nothing

	  HasAccess = iReturnValue

End Function


'-------------------------------------------------------------------------------------------------------
' Sub ShowFoldersAndDocs( iOrgId, iFolderId )
'-------------------------------------------------------------------------------------------------------
Function GetFolderId( iOrgId, sFolder )
	Dim sSql, oRs

	sSql = "SELECT folderid FROM DocumentFolders WHERE orgid = " & iOrgId & " AND FolderName = '" & sFolder & "'"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetFolderId = oRs("folderid")
	Else
		GetFolderId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 



%>