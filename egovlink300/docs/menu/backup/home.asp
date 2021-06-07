<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="../../includes/start_modules.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: home.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Documents.
'
' MODIFICATION HISTORY
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

Session("RedirectPage") = Request.ServerVariables("SCRIPT_NAME") 
Session("RedirectLang") = "Return to Documents"

%>

<html>
<head>
  <title>E-Gov Services - <%=sOrgName%></title>

  <!-- <link rel="stylesheet" type="text/css" href="menu.css" /> -->
  <link rel="stylesheet" type="text/css" href="../../css/styles.css" />
  <link rel="stylesheet" type="text/css" href="../../global.css" />
  <link rel="stylesheet" type="text/css" href="../../css/style_<%=iorgid%>.css" />


</head>

<!--#Include file="../../include_top.asp"-->

  <script language="JavaScript1.2" src="../../scripts/doctreenav.js"></script>
  <script language="JavaScript1.2" src="../../scripts/docfolderopen.js" ></script>

<!--BODY CONTENT-->


<tr><td valign="top">
<p>
	<font class="pagetitle"><%=sOrgName%> Online Documents</font><br />
</p>

	<!--BEGIN:  USER REGISTRATION - USER MENU-->
	<%	RegisteredUserDisplay( "../../" ) %>
	<!--END:  USER REGISTRATION - USER MENU-->




<!--BEGIN: ADOBE CODE-->
	<div style="width:750px; padding-left:10px;padding-bottom:10px;">Some of the pages within this section link to Portable Document Format (PDF) files which require a PDF reader to view. You may download a free copy of Adobe&reg; Reader&reg; if you do not already have it on your computer.<br><br>
	<A href='http://www.adobe.com/products/acrobat/readstep2.html' target='_blank' title='Get Adobe Acrobat Reader Plug-in Here'><img border=0 src="../../images/adreader.gif" hspace=10>Get Adobe Reader.</a>
	</div>
<!--END: ADOBE CODE-->


  <table border=0>
  <tr>
  <td valign=top style="padding-left:10px;">
 <form action="../search/search.asp" method=post id=form1 name=frmSearch>
	<p><font class=searchlabel>Search Documents:</font><br>
	          <input type=hidden name="Action" value="Go">
	          <input type="text" name="SearchString" style="background-color:#eeeeee; border:1px solid #000000; width:144px;"><br>
	          <div class="quicklink" align="right"><a href="#" onClick='document.frmSearch.submit()'><img src="images/go.gif" border="0"><font class=searchlink>Search<font class=searchlabel></font></a>&nbsp;&nbsp;</div>
	        
	</p>
	</form>
  </td>
  <td valign=top>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </td>
  <td valign=top>
					<!--BEGIN: DISPLAY DOCUMENT/FOLDER TREE-->
					<%
					' GENERATE DOCUMENT/FOLDER TREE
					
					' GET CITY NAME
					sLocationName =  GetVirtualDirectyName()

					Set FSO = CreateObject("Scripting.FileSystemObject")

					' BEGIN FOLDER TREE
					response.write "<ul>" & vbcrlf

						' LIST ALL FOLDERS AND DOCUMENTS
						ENUMERATEDOCUMENTS FSO.GetFolder(Server.Mappath("/public_documents300/" & sLocationName & "/published_documents/")),Server.Mappath("/public_documents300/" & sLocationName & "/published_documents/"),"public_documents300/" & sLocationName & "/published_documents"  

					response.write "</ul>"
					%>


					<!--BEGIN: CODE TO OPEN SELECTED FOLDER FROM EXTERNAL LINK IF FOLDER ID SUPPLIED-->
					<% If request("path") <> "" Then %>
						<script language="JavaScript">
							<% 
								' OPEN FOLDER PATH SPECIFIED
								SubOpenFolderTree(replace(request.querystring("path"),"/custom/pub","")) 
							
							%>
						</script>
					<% End If %>
					<!--END: CODE TO OPEN SELECTED FOLDER FROM EXTERNAL LINK-->


					<% ' DESTROY OBJECTS
					   SET FSO = NOTHING
					%>

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
	Call LogPageVisit(iSectionID,sDocumentTitle,sURL,datDate,datDateTime,sVisitorIP,iorgid)
'--------------------------------------------------------------------------------------------------
' END: VISITOR TRACKING
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------

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
						'SET MARGIN TO -20 TO ALIGN SUB FOLDERS AND DOCUMENTS WITHIN A FOLDER PROPERLY.
						'response.write  "<li style=""margin-left:-20px;"" ><a class=""documentlist"" TARGET=""DOCUMENTS"" href=""" & sHyperlink  & """ ><img src=""" & GetFileIcon( File.Name ) & """ border=""0"" /> " & UCASE(File.Name) & " ( " & sFileSize & ") </a></li>"  & vbcrlf
						response.write  "<li class=""" & GetListItemClassName( File.Name ) & """><a class=""documentlist"" TARGET=""DOCUMENTS"" href=""" & sHyperlink  & """ > " & UCASE(File.Name) & " ( " & sFileSize & ") </a></li>"  & vbcrlf
						' <img src=""" & GetFileIcon( File.Name ) & """ border=""0"" />
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
' FUNCTION GETFILEICON(SNAME)
'-------------------------------------------------------------------------------------------------------
Function GetFileIcon(sName)

	sReturnValue = "images/msie.gif"

	Select Case lcase(right(sName,3))
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
		Case "gif"
            sReturnValue = "images/imageicon.gif"
        Case "jpg"
            sReturnValue = "images/imageicon.gif"
	End Select

	  GetFileIcon = sReturnValue

End Function


'-------------------------------------------------------------------------------------------------------
' Function GetListItemClassName( sName )
'-------------------------------------------------------------------------------------------------------
Function GetListItemClassName( sName )

	sReturnValue = "msie"

	Select Case LCase(Right(sName,3))
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
' FUNCTION GETFOLDERPOSITION(FOLDER,IFOLDERCOUNT)
'-------------------------------------------------------------------------------------------------------
Function GetFolderPosition(Folder,FindFolder,iFolderCount,iPosition)

		iStart = 0

		' GET SELECTED FOLDER NAME
			iStart = instrrev(request("path"),"/")
			If iStart <> 0 Then
				sFolderName = RIGHT(request("path"),Len(request("path")) - iStart)
			End If

		
		' LIST CONTENTS OF FOLDER (SUBFOLDERS AND FILES)
		For Each SubFolder in Folder.SubFolders

			If UCASE(SubFolder.Path) = UCASE(FindFolder) Then
				iPosition = iFolderCount
				GetFolderPosition = iPosition
			End If

			iFolderCount = iFolderCount + 1

			' RECURSIVE CALL TO GET SUBFOLDER FOR CURRENT FOLDER
			GetFolderPosition SubFolder,FindFolder,iFolderCount,iPosition

		Next

		' RETURN FOLDER POSITION
		GetFolderPosition = iPosition

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
' SUB SUBOPENFOLDERTREE(SPATH)
'-------------------------------------------------------------------------------------------------------
Sub SubOpenFolderTree( sPath )

	' BUILD ARRAY OF FOLDERS FROM PATH VALUE
	arrFolders = split(sPath,"/")

	' LOOP THRU EACH FOLDER IN PATH STARTING AT CITY HOME TO OPEN SPECIFIED FOLDER
	For iFolder = 4 to UBOUND(arrFolders)
		
		' BUILD VIRTUAL PATH TO FOLDER
		kFolder = ""
		For jFolder = 1 to iFolder
			kFolder =  kFolder &  "/" & arrFolders(jFolder) 
		Next 

		' GET CURRENT FOLDER LOCATION IN FOLDER UL TREE
		iPos = GetFolderPosition(FSO.GetFolder(Server.Mappath("/public_documents300/" & sLocationName & "/published_documents/")),Server.Mappath(kFolder),0,-1)
		
		' OPEN SPECIFIED FOLDER
		response.write "var oList=document.all ? document.all[""foldinglist""] : document.getElementsByName('foldinglist');"
		response.write "oList[" & iPos & "].style.display='';"

	Next


End Sub

%>
