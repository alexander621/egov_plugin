<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="../../includes/start_modules.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: home_ada.asp
' AUTHOR: Steve Loar
' CREATED: 8/17/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Documents - This is the old way of displaying documents but with everything showing
'
' MODIFICATION HISTORY
' 1.5	8/16/2010	Steve Loar - Initial version from old code
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim FSO, sLocationName, oDocsOrgs

'response.redirect "../../outage.html"

Set oDocsOrg = New classOrganization

%>

<html>
<head>
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
	<title>E-Gov Services - <%=sOrgName%></title>

	<link rel="stylesheet" type="text/css" href="../../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../../global.css" />
	<link rel="stylesheet" type="text/css" href="../../css/style_<%=iorgid%>.css" />
	<link rel="stylesheet" type="text/css" href="docs_ada.css" />

</head>

<!--#Include file="../../include_top.asp"-->

<!--BODY CONTENT-->

	<tr><td valign="top">
	<p>
		<!--BEGIN:  RSS FEED Picks-->
		<table border="0" cellspacing="0" cellpadding="0" style="max-width:800px;padding-left:10px;">
			<tr>
				<td>
      <%
       'Build the welcome message
        lcl_org_name        = oDocsOrg.GetOrgName()
        lcl_org_state       = oDocsOrg.GetState()
        lcl_org_featurename = "Online Documents"

        oDocsOrg.buildWelcomeMessage iorgid, lcl_orghasdisplay_action_page_title, lcl_org_name, lcl_org_state, lcl_org_featurename
		Set oDocsOrg = Nothing 

      %>
					<!--<font class="pagetitle"><%'sOrgName%> Online Documents</font>-->
					<% checkForRSSFeed iOrgId, "", "", "DOCS", sEgovWebsiteURL %>
					<!-- BEGIN:  USER REGISTRATION - USER MENU -->
					<%	RegisteredUserDisplay( "../../" ) %>
					<!--END:  USER REGISTRATION - USER MENU-->
				</td>
			</tr>
		</table>
		<!--END:  RSS FEED Picks-->
	</p>

	<!--BEGIN: ADOBE CODE-->
	<table border="0" cellspacing="0" cellpadding="0" style="max-width:750px;padding-left:10px; padding-bottom:10px;">
	  <tr>
		  <td>
			  Some of the pages within this section link to Portable Document Format (PDF) files which require a PDF reader to view. 
			  You may download a free copy of Adobe&reg; Reader&reg; if you do not already have it on your computer.<br /><br />
				<a href='http://www.adobe.com/products/acrobat/readstep2.html' target='_blank' title='Get Adobe Acrobat Reader Plug-in Here'><img border="0" src="../../images/adreader.gif" hspace="10" />Get Adobe Reader.</a>
		  </td>
	  </tr>
	</table>
	<!--END: ADOBE CODE-->
 
	<table border="0" cellpading="0" cellspacing="0" class="respTable">
		<tr>
		<td valign="top" style="padding-left:10px;">
			<!--form action="../search/search.asp" method="post" id="form1" name="frmSearch">
				<p><font class=searchlabel>Search Documents:</font><br />		
					<input type=hidden name="Action" value="Go">
					<input type="text" name="SearchString" style="background-color:#eeeeee; border:1px solid #000000; width:144px;"><br />
					<div id="docsearch" align="right">
						<input type="submit" class="button" value="Search" /><br />
					</div>
				</p>
			</form-->
			<form action="../archive.asp?docsMonth=<%=Month(Now)%>&docsYear=<%=Year(Now)%>" method="post">
				<p>
					<input type="submit" class="button" value="Recently Uploaded Docs" />
				</p>
			</form>
			<div id="docswitch">
				To view this page with JavaScript enabled <a href="home.asp">Click Here</a>. 
			</div>
		</td>
		<td valign="top">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		</td>
		<td valign="top">
			<!--BEGIN: DISPLAY DOCUMENT/FOLDER TREE-->
			<div class="treeContainer">
			<%
			' GENERATE DOCUMENT/FOLDER TREE
			sLocationName =  GetVirtualDirectyName()

			Set FSO = CreateObject("Scripting.FileSystemObject")

			response.write vbcrlf & "<ul>"

			on error resume next
			ENUMERATEDOCUMENTS FSO.GetFolder(Server.Mappath("/public_documents300/" & sLocationName & "/published_documents/")), Server.Mappath("/public_documents300/" & sLocationName & "/published_documents/"), "public_documents300/" & sLocationName & "/published_documents"  
			on error goto 0

			response.write vbcrlf & "</uL>"
			%>

			<!--BEGIN: CODE TO OPEN SELECTED FOLDER FROM EXTERNAL LINK IF FOLDER ID SUPPLIED-->
			<% If request("path") <> "" Then %>
				<script language="JavaScript1.2">
					//OpenFolder(<%'=clng(GetFolderPosition(FSO.GetFolder(Server.Mappath("/public_documents300/" & sLocationName & "/published_documents/")),0,False,-1))%>,<%=clng(GetFolderPosition(FSO.GetFolder(Server.Mappath("/public_documents300/" & sLocationName & "/published_documents/")),0,True,-1))%>);
				</script>
			<% End If %>
			<!--END: CODE TO OPEN SELECTED FOLDER FROM EXTERNAL LINK-->


			<% ' DESTROY OBJECTS
			SET FSO = NOTHING				
			%>

			</div>
			<!--END: DISPLAY DOCUMENT/FOLDER TREE-->
		</td>
		</tr>
	</table>


<!--#Include file="../../include_bottom.asp"--> 


<%
'--------------------------------------------------------------------------------------------------
' BEGIN: VISITOR TRACKING
'--------------------------------------------------------------------------------------------------
	iSectionID = 4
	sDocumentTitle = "MAIN"
	sURL = request.servervariables("SERVER_NAME") &":/" & request.servervariables("URL") & "?" & request.servervariables("QUERY_STRING")
	datDate = Date()	
	datDateTime = Now()
	sVisitorIP = request.servervariables("REMOTE_ADDR")
	LogPageVisit iSectionID, sDocumentTitle, sURL, datDate, datDateTime, sVisitorIP, iorgid
'--------------------------------------------------------------------------------------------------
' END: VISITOR TRACKING
'--------------------------------------------------------------------------------------------------
%>


<%
'--------------------------------------------------------------------------------------------------
' Functions and Subroutines
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------------
' void ENUMERATEDOCUMENTS(FOLDER)
'-------------------------------------------------------------------------------------------------------
Sub ENUMERATEDOCUMENTS( ByVal Folder, ByVal sPath, ByVal sVpath )
	Dim sVirtualpath, sTempPath, SubFolder
     
	' BUILD HYPERLINK BASE PATH
	sVirtualpath = Replace(Folder.Path,sPath,"")
	sTempPath =  Replace(Folder.Path,sPath,"")
	sVirtualpath = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & sVpath & Replace(sVirtualpath,"\","/") 

	' LIST CONTENTS OF FOLDER (SUBFOLDERS AND FILES)
	For Each SubFolder In Folder.SubFolders
		
		sVirtualpath2 =  Replace(sVpath,"public_documents300","/public_documents300/custom/pub") & Replace(sTempPath,"\","/") & "/" & SubFolder.Name
		
		' CHECK SECURITY FOR ACCESS
		If HasAccess(iorgid,request.Cookies("userid"),sVirtualpath2) Then

			' WRITE FOLDER INFORMATION
			response.write vbcrlf & "<li class=""foldheader"" >" & SubFolder.Name & "</li>"
			'response.write vbcrlf & "<ul id=""foldinglist"" style=""display:none;"">"
			response.write vbcrlf & "<ul class=""filelist"">"
		
			' RECURSIVE CALL TO GET ANY SUBFOLDERS OF THE CURRENT FOLDER
			ENUMERATEDOCUMENTS Subfolder, sPath, sVpath
			
			' LIST FILES IN THE CURRENT FOLDER
			For Each File In Subfolder.Files
				
				' GET FILE SIZE
				If File.Size > 1024 Then
					sFileSize = FormatNumber((File.Size / 1024),0)  & " KB"
				Else
					sFileSize =  FormatNumber(File.Size,0) & " Bytes"
				End If
				
				sHyperlink = sVirtualPath & "/" & Subfolder.Name & "/" & File.Name
				response.write vbcrlf & "<li>"
				'response.write "<img src=""" & GetFileIcon( File.Name ) & """ />"
				'response.write "<span class=""fileimg""><img src=""images/vline.gif"" alt="""" height=""22"" width=""16""border=""0"" /></span>"
				response.write "<a class=""documentlist"" target=""DOCUMENTS"" href=""" & sHyperlink  & """ > " & UCase(File.Name) & " ( " & sFileSize & ") </a></li>" 
			Next 

			' IF FOLDER CONTAINS NO FILES OR FOLDERS SHOW AS EMPTY
			If (Subfolder.Files.Count < 1) And (Subfolder.SubFolders.Count < 1) Then
				response.write vbcrlf & "<li class=""emptyfolder""> (<i> EMPTY </i>) </li>" 
			End IF

			' CLOSE FOLDER LIST TAG
			response.write vbcrlf & "</ul>"

		End If
	Next

End Sub


'-------------------------------------------------------------------------------------------------------
' string GetFileIcon( sName )
'-------------------------------------------------------------------------------------------------------
Function GetFileIcon( ByVal sName )
	Dim sReturnValue

	sReturnValue = "images/msie.gif"

	Select Case LCase(Right(sName,3))
		Case "doc"
            sReturnValue = "images/msword.gif"
        Case "xls"
            sReturnValue = "images/msexcel.gif"
        Case "ppt"
            sReturnValue = "images/msppt.gif"
        Case "htm"
            sReturnValue = "images/msie.gif"
        Case "pdf"
            sReturnValue = "images/pdf.gif"
	End Select

	  GetFileIcon = sReturnValue

End Function


'-------------------------------------------------------------------------------------------------------
' integer GetFolderPosition( Folder, iFolderCount, blnIsParent, iPosition )
'-------------------------------------------------------------------------------------------------------
Function GetFolderPosition( ByVal Folder, ByVal iFolderCount, ByVal blnIsParent, ByVal iPosition )
	Dim sTemp, iStart, sFolderName, SubFolder

	iStart = 0

	' GET PARENT FOLDER NAME
	If blnIsParent Then
		iStart = InStr(request("path"),"/published_documents/")
		If iStart <> 0 Then
			sTemp = Replace(request("path"), Left(request("path"), iStart+20),"")
			If InStr(sTemp,"/") <> 0 Then
				sFolderName = Left(sTemp,InStr(sTemp,"/")-1)
			Else
				sFolderName = sTemp
			End If
		End If
	Else
	' GET SELECTED FOLDER NAME
		iStart = InstrRev(request("path"),"/")
		If iStart <> 0 Then
			sFolderName = RIGHT(request("path"),Len(request("path")) - iStart)
		End If
	End If

	' LIST CONTENTS OF FOLDER (SUBFOLDERS AND FILES)
	For Each SubFolder in Folder.SubFolders
		If UCase(sFolderName) = UCase(SubFolder.Name) Then
			iPosition = iFolderCount
		End If

		iFolderCount = iFolderCount + 1

		' RECURSIVE CALL TO GET SUBFOLDER FOR CURRENT FOLDER
		GetFolderPosition SubFolder, iFolderCount, blnIsParent, iPosition  

	Next

	' RETURN FOLDER POSITION
	GetFolderPosition = iPosition

End Function


'-------------------------------------------------------------------------------------------------------
' boolean HasAccess( iorgid, iuserid, strvpath )
'-------------------------------------------------------------------------------------------------------
Function HasAccess( ByVal iorgid, ByVal iuserid, ByVal strvpath )
	Dim oCnn, oRs, sSql, iReturnValue
  
	On Error Resume Next

	iReturnValue = False

	Set oCnn = Server.CreateObject("ADODB.Connection")
	oCnn.Open Application("DSN")
	sSql = "EXEC CHECKFOLDERACCESS '" & iorgid & "','" & iuserid & "','" & strvpath & "'"
	Set oRs = oCnn.Execute(sSql)

	If Not oRs.EOF Then
		If oRs("folderid") >= 0 Then
			iReturnValue = True
		End If 
	End If

	oRs.Close
	Set oRs = Nothing
	Set oCnn = Nothing

	HasAccess = iReturnValue

End Function




%>
