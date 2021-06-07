<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="../../includes/start_modules.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: home.asp
' AUTHOR: Steve Loar	
' CREATED: 3/26/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Documents.
'
' MODIFICATION HISTORY
' 1.5	8/16/2010	Steve Loar - Cleaned up code in prep for ADA copy or changes.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 dim iFolderCount, aPath, bHavePath, sPath, sLocationName, oDocsOrg, lcl_org_name, lcl_org_state
 dim FSO, lcl_org_featurename

 if not OrgHasFeature(iorgid,"documents") then response.redirect "../../default.asp"
 'response.redirect "../../outage.html"

 if request("path") <> "" then
	   if InStr(UCase(request("path")),"DECLARE") > 0 And InStr(UCase(request("path")),"VARCHAR") > 0 And InStr(UCase(request("path")),"EXEC") > 0 then
     		response.redirect "home.asp"
   	end if

   	if InStr(UCase(request("path")),"SELECT") > 0 And InStr(UCase(request("path")),"VARCHAR") > 0 And InStr(UCase(request("path")),"CAST") > 0 then
     		response.redirect "home.asp"
   	end if
 end if

	sPath = ""

 if request("path") <> "" then
  	 'path=/public_documents300/eclink/published_documents/City+Council/Council%20Minutes/2008
   	'response.write "/public_documents300/" & GetVirtualDirectyName() & "/published_documents/" & "<br />"
   	'response.write sPath & "<br />"
   	sPath = Replace(Track_DBSafe(request("path")),"/custom/pub","")
   	sPath = Replace( sPath, "/public_documents300/" & GetVirtualDirectyName() & "/published_documents/", "" )
 end if

 session("RedirectPage") = Request.ServerVariables("SCRIPT_NAME") 
 session("RedirectLang") = "Return to Documents"

 set oDocsOrg = New classOrganization

'Check for a Google Custom Search Engine ID
 lcl_googleSearchID = getGoogleSearchID(iOrgID, "googlesearchid_documents")
%>
<html>
<head>
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
	 <title>E-Gov Services - <%=sOrgName%></title>

 	<link rel="stylesheet" type="text/css" href="../../yui/build/tabview/assets/skins/sam/tabview.css" />
 	<link rel="stylesheet" type="text/css" href="../../yui/build/treeview/assets/skins/sam/treeview.css" />
 	<link rel="stylesheet" type="text/css" href="../../yui/examples/treeview/assets/css/folders/tree.css" /> 
 	<link rel="stylesheet" type="text/css" href="../../css/styles.css" />
 	<link rel="stylesheet" type="text/css" href="../../global.css" />
 	<link rel="stylesheet" type="text/css" href="../../css/style_<%=iorgid%>.css" />

<style type="text/css">
  #searchbox {
     whitespace: nowrap;
     padding-left: 10px;
     min-width: 300px;
  }

  #documentsbox {
     width: 100%;
  }

  #googleSearch {
     width: 70%;
  }
</style>

  <script type="text/javascript" src="../../yui/build/yahoo-dom-event/yahoo-dom-event.js" ></script>
	 <script type="text/javascript" src="../../yui/build/animation/animation-min.js"></script>
	 <script type="text/javascript" src="../../yui/build/element/element-beta.js"></script>

 	<!-- TreeView source file -->
 	<script src = "../../yui/build/treeview/treeview-min.js"></script>  
 	<script type="text/javascript" src="../../prototype/prototype-1.6.0.2.js"></script>
 	<script type="text/javascript" src="../../scripts/ajaxLib.js"></script>

<script type="text/javascript">
<!--
		var tree;
		var oCurrentNode;

		function ValidateSearch()
		{
			if ($("SearchString").value == "")
			{
				$("SearchString").focus();
				alert('Please enter some text in the box before starting a search.');
				return;
			}
			document.frmSearch.submit();
		}
//-->
</script>

<% if lcl_googleSearchID <> "" then %>
<!-- Put the following javascript before the closing </head> tag. -->
<script>
//  (function() {
//    var cx = '<%'lcl_googleSearchID%>';
//    var gcse = document.createElement('script'); gcse.type = 'text/javascript'; gcse.async = true;
//    gcse.src = (document.location.protocol == 'https:' ? 'https:' : 'http:') +
//        '//www.google.com/cse/cse.js?cx=' + cx;
//    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(gcse, s);
//  })();
  
  (function() {
    var cx = '<%=lcl_googleSearchID%>';
    var gcse = document.createElement('script');
    gcse.type = 'text/javascript';
    gcse.async = true;
    gcse.src = (document.location.protocol == 'https:' ? 'https:' : 'http:') +
        '//www.google.com/cse/cse.js?cx=' + cx;
    var s = document.getElementsByTagName('script')[0];
    s.parentNode.insertBefore(gcse, s);
  })();
</script>
<% end if %>

</head>

	<!--#Include file="../../include_top.asp"-->
<%
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"">" & vbcrlf
  response.write "  	       <p>" & vbcrlf

 'BEGIN: RSS Feed Picks -------------------------------------------------------
  response.write "  	       <table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""padding-left:10px; max-width:800px;"">" & vbcrlf
  response.write "  	         <tr>" & vbcrlf
  response.write "  	       				  <td>" & vbcrlf

 'Build the welcome message
  lcl_org_name        = oDocsOrg.GetOrgName()
  lcl_org_state       = oDocsOrg.GetState()
  lcl_org_featurename = "Online Documents"
  lcl_org_featurename = oDocsOrg.GetOrgFeatureName("documents")

  oDocsOrg.buildWelcomeMessage iorgid, _
                               lcl_orghasdisplay_action_page_title, _
                               lcl_org_name, _
                               lcl_org_state, _
                               lcl_org_featurename

		set oDocsOrg = nothing 

  checkForRSSFeed iOrgId, _
                  "", _
                  "", _
                  "DOCS", _
                  sEgovWebsiteURL

  RegisteredUserDisplay( "../../" )

  response.write "  	       				  </td>" & vbcrlf
  response.write "  	       		</tr>" & vbcrlf
  response.write "  	       </table>" & vbcrlf
 'END: RSS Feed Picks ---------------------------------------------------------

  response.write "  	       </p>" & vbcrlf
	
 'BEGIN: Adobe Code -----------------------------------------------------------
  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""max-width:750px;padding-left:10px; padding-bottom:10px;"">" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    Some of the pages within this section link to Portable Document Format (PDF) files which require a PDF reader to view. " & vbcrlf
  response.write "                    You may download a free copy of Adobe&reg; Reader&reg; if you do not already have it on your computer.<br /><br />" & vbcrlf
  if iorgid = "190" then
	  response.write "As a courtesy to citizens of both communities, and because our communities share many overlapping interests, we are making available meeting minutes and various documents from the Municipality of Chimney Rock Village. For more information about Chimney Rock Village, visit <a href=""http://www.chimneyrockvillage.com"">www.chimneyrockvillage.com</a><br /><br />"
  end if
  if iorgid = "186" then
	  response.write "Fillable forms must be downloaded to your computer to be completed. These forms will not work with certain browsers, including Google Chrome.<br /><br />"
  end if
  if iorgid = "115" then
  %>
<b>Please enter one or all of the criteria in the search box as specified below:</b><br />
<br />
Date of Accident &bull; mmddyyyy (ex. 09232017) <br />
Street Name &bull; (ex. Branch Hill Guinea) <br />
Party Involved &bull; (ex. Smith)<br />
<br />
<b>Example #1:</b> 09232017<br />
<b>Example #2:</b> 09232017 Branch Hill Guinea<br />
<b>Example #3:</b> 09232017 Branch Hill Guinea Smith<br />
<br />
<b>*** PLEASE NOTE: Crash reports are only kept online for 60 days ***</b>
<br />
If you are not able to find the report you are looking for, call the MTPD at (513) 248.3721 <br />
<br />
  
  <%
  end if
  response.write "                    <a href='http://www.adobe.com/products/acrobat/readstep2.html' target='_blank' title='Get Adobe Acrobat Reader Plug-in Here'><img border=""0"" src=""../../images/adreader.gif"" hspace=""10"" />Get Adobe Reader.</a>" & vbcrlf
  response.write "  	       		    </td>" & vbcrlf
  response.write "  	       		</tr>" & vbcrlf
  response.write "  	       </table>" & vbcrlf
 'END: Adobe Code -------------------------------------------------------------

 'BEGIN: Google Search --------------------------------------------------------
  if lcl_googleSearchID <> "" then
    'Place this tag where you want the search box to render
    'response.write "<p><gcse:searchbox-only></gcse:searchbox-only></p>" & vbcrlf
     response.write "<div id=""googleSearch"">" & vbcrlf
     response.write "  <gcse:search></gcse:search>" & vbrlf
     response.write "</div>" & vbcrlf
  end if
 'END: Google Search ----------------------------------------------------------

  response.write "  	       <table border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbcrlf
  response.write "  	         <tr>" & vbcrlf
  response.write "  	       	     <td valign=""top"" id=""searchbox"">" & vbcrlf

  if lcl_googleSearchID = "" and 1=2 then
     response.write "<p>" & vbcrlf
     response.write "<form action=""../search/search.asp"" method=""post"" id=""form1"" name=""frmSearch"">" & vbcrlf
			  response.write "  <input type=""hidden"" name=""Action"" value=""Go"" />" & vbcrlf
     response.write "<font class=""searchlabel"">Search Documents:</font><br />" & vbcrlf
     response.write "<input type=""text"" id=""SearchString"" name=""SearchString"" size=""65"" maxlength=""100"" style=""background-color:#eeeeee; border:1px solid #000000; width:144px;"" /><br />" & vbcrlf
     response.write "<div class=""quicklink"" align=""right"">" & vbcrlf
     response.write "  <a href=""#"" onClick=""ValidateSearch();""><img src=""images/go.gif"" border=""0"" /><font class=""searchlink"">Search</font></a>&nbsp;&nbsp;" & vbcrlf
     response.write "</div>" & vbcrlf
     response.write "</form>" & vbcrlf
     response.write "</p>" & vbcrlf
  end if

  response.write "  	       	     			<p>" & vbcrlf
  response.write "  	       	     			  <input type=""button"" name=""recentDocsButton"" id=""recentDocsButton"" class=""button"" value=""Recently Uploaded Docs"" onclick=""location.href='../archive.asp?docsMonth=" & month(now) & "&docsYear=" & year(now) & "'"" />" & vbcrlf
  response.write "  	       	     			</p>" & vbcrlf
  response.write "  	       	     		 <div id=""docswitch"">" & vbcrlf
  response.write "  	       	     			  To view this page wi thout JavaScript <a href=""home_ada.asp"">Click Here</a>." & vbcrlf
  response.write "  	       	     		 </div>" & vbcrlf
  response.write "  	       	     </td>" & vbcrlf
  response.write "  	       	     <td valign=""top"" id=""documentsbox"">" & vbcrlf

 'BEGIN: Display Document/Folder Tree -----------------------------------------
 'GET CITY NAME
 	sLocationName =  GetVirtualDirectyName()
 	Set FSO = CreateObject("Scripting.FileSystemObject")

	'BEGIN: Folder Tree ----------------------------------------------------------
  response.write "  	       	         <div id=""treeDiv1"" class=""treeContainer"">" & vbcrlf
  response.write "  	       	           <ul>" & vbcrlf

	'List all folders and documents
	on error resume next
 	ShowRootFolders FSO.GetFolder(Server.Mappath("/public_documents300/" & sLocationName & "/published_documents/")), _
                  Server.Mappath("/public_documents300/" & sLocationName & "/published_documents/"), _
                  "public_documents300/" & sLocationName & "/published_documents"
	on error goto 0

  response.write "  	       	           </ul>" & vbcrlf
  response.write "  	       	         </div>" & vbcrlf

 	set FSO = nothing 
 'END: Display Document/Folder Tree -------------------------------------------
  response.write "  	       	     </td>" & vbcrlf
  response.write "  	       	 </tr>" & vbcrlf
  response.write "  	       </table>" & vbcrlf
%>
	<script>
	<!--
		
		treeInit();

<%		if sPath <> "" then	%>
			YAHOO.util.Event.onDOMReady(treeReady);
<%		end if %>

		function treeReady()
		{
			var rootNode = tree.getRoot();
			var sPath = '<%=sPath%>';
			var aPath = sPath.split("/");

			ShowPath(rootNode, aPath, 0);
		}

		function ShowPath(oNode, aPath, x)
		{
			if (x < aPath.length)
			{
				sFolder = aPath[x];
				for (var i=0; i < oNode.children.length; i++)
				{
					if (oNode.children[i].data.label == sFolder)
					{
						oNode.children[i].expand();
						ShowPath( oNode.children[i], aPath, x + 1);
						break;
					}
				}
			}
		}

		function treeInit() 
		{ 
			tree = new YAHOO.widget.TreeView("treeDiv1");   
			tree.setDynamicLoad(loadNodeData);  
			tree.render();  
		}   

		function OpenDocFolder( oNode, sFolder )
		{
			for (var i=0,len=oNode.children.length;i<len;++i) 
			{
				if (oNode.children[i].data.label == sFolder)
				{
					oNode.children[i].focus();  
					//oNode.children[i].expand;
					//alert('found');
					break;
				}
			}
		}

		function loadNodeData(oNode, fnLoadComplete)  
		{
			var sendpath = oNode.data.label;
			oCurrentNode = oNode;

			var oNew = oNode.parent;
			if (oNew.data)
			{
				while (oNew.data)
				{
					sendpath = oNew.data.label + '/' + sendpath;
					oNew = oNew.parent;
				}
			}

			doAjax('jsonarr.asp', 'path=' + sendpath, 'AddLeafNodes', 'get', '0');
		}

		function AddLeafNodes( oResponse )
		{
			var json = oResponse.evalJSON(true);

			var node1;
			if (json[0][0] == 'EMPTY')
			{
				node1 = new YAHOO.widget.HTMLNode('(<i>EMPTY</i>)',oCurrentNode);
				node1.isLeaf = true;
			}
			else
			{
				if (json[0][0] != 'NOFILE')
				{
					for (var i = 0; i < json.length; i++)
					{
						node1 = new YAHOO.widget.HTMLNode('<a class="documentlist" TARGET="DOCUMENTS" href="' + json[i][0] + '">' + json[i][1] + "</a>",oCurrentNode);
						node1.isLeaf = true;
					}
				}
			}

			oCurrentNode.loadComplete( ); 
		}

		var onLabelClick = function(oNode) 
		{ 
			alert(oNode.data.label);
		}

	//-->
	</script>

<%
'BEGIN: Visitor Tracking ------------------------------------------------------
	iSectionID     = 4
	sDocumentTitle = "MAIN"
	sURL           = request.servervariables("SERVER_NAME") &":/" & request.servervariables("URL") & "?" & request.servervariables("QUERY_STRING")
	datDate        = Date()
	datDateTime    = Now()
	sVisitorIP     = request.servervariables("REMOTE_ADDR")

	LogPageVisit iSectionID, _
              sDocumentTitle, _
              sURL, _
              datDate, _
              datDateTime, _
              sVisitorIP, _
              iorgid
'END: Visitor Tracking --------------------------------------------------------
%>
<!--#Include file="../../include_bottom.asp"-->
<%
'------------------------------------------------------------------------------
Sub ShowRootFolders( ByVal Folder, ByVal sPath, ByVal sVpath )
	Dim sVirtualpath, sTempPath, SubFolder, sVirtualpath2

	' BUILD HYPERLINK BASE PATH
	sVirtualpath = Replace(Folder.Path,sPath,"")
	sTempPath =  Replace(Folder.Path,sPath,"")
	sVirtualpath = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & sVpath & Replace(sVirtualpath,"\","/")

	' LIST CONTENTS OF FOLDER (SUBFOLDERS AND FILES)
	For Each SubFolder in Folder.SubFolders

		sVirtualpath2 =  Replace(sVpath,"public_documents300","/public_documents300/custom/pub") & Replace(sTempPath,"\","/") & "/" & SubFolder.Name

		' CHECK SECURITY FOR ACCESS
		If HasAccess( iOrgId, request.Cookies("userid"), sVirtualpath2 ) Then

			' WRITE FOLDER INFORMATION
			response.write vbcrlf & "<li id=""foldheader"">" & SubFolder.Name & "</li>"
			response.write vbcrlf & "<ul id=""foldinglist"" name=""foldinglist"" style=""display:none;"" >"

			ShowRootFolders Subfolder, sPath, sVpath

			' CLOSE FOLDER LIST TAG
			response.write vbcrlf & "</ul>"
		End If
	Next

End Sub


'-----------------------------------------------------------------------------------
' void ENUMERATEDOCUMENTS( Folder, sPath, sVpath )
'-----------------------------------------------------------------------------------
Sub ENUMERATEDOCUMENTS( ByVal Folder, ByVal sPath, ByVal sVpath )
	Dim sVirtualpath, sTempPath, SubFolder, sVirtualpath2, File, sFileSize, sHyperlink

	' BUILD HYPERLINK BASE PATH
	sVirtualpath = Replace(Folder.Path,sPath,"")
	sTempPath =  Replace(Folder.Path,sPath,"")
	sVirtualpath = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & sVpath & Replace(sVirtualpath,"\","/")

	' LIST CONTENTS OF FOLDER (SUBFOLDERS AND FILES)
	For Each SubFolder in Folder.SubFolders

		sVirtualpath2 =  Replace(sVpath,"public_documents300","/public_documents300/custom/pub") & Replace(sTempPath,"\","/") & "/" & SubFolder.Name

		' CHECK SECURITY FOR ACCESS
		If HasAccess( iorgid, request.Cookies("userid"), sVirtualpath2 ) Then

			' WRITE FOLDER INFORMATION
			response.write vbcrlf & "<li id=""foldheader""> " & SubFolder.Name & "</li>"
			response.write vbcrlf & "<ul id=""foldinglist"" name=""foldinglist"" style=""display:none;"" >"

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
				response.write vbcrlf & "<li><a class=""documentlist"" TARGET=""DOCUMENTS"" href=""" & sHyperlink  & """ > " & UCase(File.Name) & " ( " & sFileSize & ") </a></li>"
			Next

			' IF FOLDER CONTAINS NO FILES OR FOLDERS SHOW AS EMPTY
			If (Subfolder.Files.Count < 1) AND (Subfolder.SubFolders.Count < 1) Then
				'SET MARGIN TO -20 TO ALIGN SUB FOLDERS AND DOCUMENTS WITHIN A FOLDER PROPERLY.
				response.write vbcrlf & "<li class=""emptyfolder"" STYLE=""margin-left:-20px;""> (<i> EMPTY </i>) </li>" 
			End IF

			' CLOSE FOLDER LIST TAG
			response.write vbcrlf & "</ul>"
		Else 
			' Added to allow direct folder links to work while skipping restricted access folders
			response.write vbcrlf & "<li id=""foldheader"" style=""display:none;"" ></li>"
			response.write vbcrlf & "<ul id=""foldinglist"" name=""foldinglist"" style=""display:none;"" ></ul>"
			
			' RECURSIVE CALL TO GET ANY SUBFOLDERS OF THE CURRENT FOLDER
			ENUMERATEDOCUMENTS Subfolder, sPath, sVpath
		End If
	Next

End Sub


'-----------------------------------------------------------------------------------
' boolean UserHasDocFolderAccess( iOrgId, iUserId, iFolderId )
'-----------------------------------------------------------------------------------
Function UserHasDocFolderAccess( ByVal iOrgId, ByVal iUserId, ByVal iFolderId )
	Dim oRs, sSql, iReturnValue

	On Error Resume Next

	iReturnValue = False

	Set oCnn = Server.CreateObject("ADODB.Connection")
	oCnn.Open Application("DSN")
	sSql = "EXEC CheckFolderIdAccess '" & iOrgId & "','" & iUserId & "','" & iFolderId & "'"

	'response.write sSql & "<br /><br />"
	Set oRs = oCnn.Execute(sSql)

	If Not oRs.EOF Then
		If oRs("folderid") = iFolderId Then
			iReturnValue = True
		End If
	End If

	oRs.Close
	Set oRs = Nothing
	Set oCnn = Nothing

	UserHasDocFolderAccess = iReturnValue

End Function


'-----------------------------------------------------------------------------------
' string GetListItemClassName( sName )
'-----------------------------------------------------------------------------------
Function GetListItemClassName( ByVal sName )
	Dim sReturnValue

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


'-----------------------------------------------------------------------------------
' string GetFileImageURL( sFileName )
'-----------------------------------------------------------------------------------
Function GetFileImageURL( ByVal sFileName )
	Dim sReturnValue

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


'-----------------------------------------------------------------------------------
' boolean HasAccess( iOrgId, iUserId, sVpath )
'-----------------------------------------------------------------------------------
Function HasAccess( ByVal iOrgId, ByVal iUserId, ByVal sVpath )
	Dim sSql, oRs, oCnn, iReturnValue

	  On Error Resume Next

	  iReturnValue = False

	  Set oCnn = Server.CreateObject("ADODB.Connection")
	  oCnn.Open Application("DSN")

	  sSql = "EXEC CHECKFOLDERACCESS '" & iOrgId & "','" & iUserId & "','" & sVpath & "'"
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


'-----------------------------------------------------------------------------------
' int ShowFoldersAndDocs( iOrgId, iFolderId )
'-----------------------------------------------------------------------------------
Function GetFolderId( ByVal iOrgId, ByVal sFolder )
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
