<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: default.asp
' AUTHOR: Steve Loar
' CREATED: 08/23/2010
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Document folder structure without frameset
'
' MODIFICATION HISTORY
' 1.0   08/24/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sLocationName, iTopCount, sRootURL, sPublished, sUnpublished, sSuccessFlag, sFolderName, sFileName
'response.redirect "../outage.html"

sLevel = "../" ' Override of value from common.asp

PageDisplayCheck "manage documents", sLevel	' In common.asp

sRootURL = "/public_documents300/custom/pub/" & session("virtualdirectory")

sPublished = sRootURL & "/published_documents/"
sUnpublished = sRootURL & "/unpublished_documents/"

sSuccessFlag = request("sf")
If sSuccessFlag <> "" Then 
	sFolderName = request("foldername")
	sFileName = request("filename")
	sNewName = request("newname")
	If sSuccessFlag = "fod" Then
		sLoadMsg = "displayScreenMsg('The folder, " & sFolderName & ", has been successfully deleted.');"
	End If 
	If sSuccessFlag = "fd" Then
		sLoadMsg = "displayScreenMsg('The document, " & sFileName & ", has been successfully deleted.');"
	End If 
	If sSuccessFlag = "nf" Then
		sLoadMsg = "displayScreenMsg('The document you attempted to delete, " & sFileName & ", does not exist. Please try again.');"
	End If 
	If sSuccessFlag = "nr" Then
		sLoadMsg = "displayScreenMsg('The document you attempted to rename, " & sFileName & ", does not exist. Please try again.');"
	End If 
	If sSuccessFlag = "fr" Then
		sLoadMsg = "displayScreenMsg('The document has been successfully renamed to " & sNewName & ".');"
	End If 
	If sSuccessFlag = "nm" Then
		sLoadMsg = "displayScreenMsg('The document you attempted to move, " & sFileName & ", does not exist. Please try again.');"
	End If 
	If sSuccessFlag = "fm" Then
		sLoadMsg = "displayScreenMsg('The document has been successfully moved.');"
	End If 
End If 

'Check for a Google Custom Search Engine ID
 lcl_googleSearchID = getGoogleSearchID(session("orgid"), "googlesearchid_documents")
%>

<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />

	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../css/jqueryFileTree.css" media="screen" />
	<link rel="stylesheet" type="text/css" href="../css/jquery.contextMenu.css" media="screen" />
	<link rel="stylesheet" type="text/css" href="docstyles.css" />

<style type="text/css">
  #content table {
     top: 0px !important;
     left: 0px !important;
  }
</style>

	<script type="text/javascript" src="../scripts/jquery-1.3.2.min.js"></script>
	<script type="text/javascript" src="../scripts/jquery.easing.1.3.js"></script>
	<script type="text/javascript" src="../scripts/jqueryFileTree.js"></script>
	<script type="text/javascript" src="../scripts/jquery.contextMenu.js"></script>
	<script type="text/javascript" src="../scripts/jquery.cookie.js"></script>

	<script type="text/javascript">
	<!--

		function loader()
		{
			<%=sLoadMsg%>
		}

		function displayScreenMsg( iMsg ) 
		{
			if(iMsg!="") 
			{
				$("#screenMsg").html("*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;");
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			$("#screenMsg").html("&nbsp;");
		}

		function ValidateSearch()
		{
			if ($("#SearchString").val() == '')
			{
				$("#SearchString").focus();
				alert("Please enter some text in the search box before starting a search.");
			}
			else
			{
				document.frmSearch.submit();
			}
		}

		function HandleFolderContextMenu(action, folderPath)
		{
			//alert(action + '\n\nFolder Path: ' + folderPath);

			if (action == 'addfoldertoroot' || action == 'addfolder')
			{
				document.frmAction.action = 'addfolder.asp';
			}

			if (action == 'deletefolder')
			{
				document.frmAction.action = 'deletefolder.asp';
			}

			if (action == 'adddocument')
			{
				document.frmAction.action = 'adddocument.asp';
			}

			if (action == 'editpublicaccess')
			{
				document.frmAction.action = 'editpublicaccess.asp';
			}

			if (action == 'editsecurity')
			{
				document.frmAction.action = 'editsecurity.asp';
			}

			$("#path").val(folderPath);
			document.frmAction.submit();
		}
		
		function HandleFileContextMenu(action, filePath)
		{
			//alert(action+'\n\n File Path : '+filePath);
			if (action == 'delete')
			{
				document.frmAction.action = 'deletefile.asp';
			}

			if (action == 'rename')
			{
				document.frmAction.action = 'renamedocument.asp';
			}

			if (action == 'move')
			{
				document.frmAction.action = 'movedocument.asp';
			}

			$("#path").val(filePath);
			document.frmAction.submit();
		}


		function showMenu()
		{
			$('.root>a').contextMenu({
			menu: 'myRootMenu'
			},
				function(action, el, pos) 
				{
					var folderPath=$(el).attr('rel');
					HandleFolderContextMenu(action, folderPath);
				}
			);

			$('.directory>a').contextMenu({
			menu: 'myFolderMenu'
			},
				function(action, el, pos) 
				{
					var folderPath=$(el).attr('rel');
					HandleFolderContextMenu(action, folderPath);
				}
			);

			$('.file>a').contextMenu({
			menu: 'myFileMenu'
			},
				function(action, el, pos) 
				{
					var filePath=$(el).attr('rel');
					HandleFileContextMenu(action, filePath);
				}
			);
			
		}
		
		$(document).ready( function() {
			
			$('#JQueryFTD_Demo').fileTree({
			  root: '/<%= Application("DocumentsRootDirectory") %>/custom/pub/<%=session("virtualdirectory")%>/',
			  script: 'jqueryFileTree2.aspx?userid=<%=session("userid")%>',
			  expandSpeed: 1000,
			  collapseSpeed: 1000,
			  multiFolder: true,
			  persist: "cookie",
			  cookieid: "JQueryFTD_Demo"
			});	
		});

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

<body onload="loader();">

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Documents</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<table id="screenMsgtable" cellpadding="0" cellspacing="0" border="0">
				<tr><td nowrap="nowrap">
					<span id="screenMsg">&nbsp;</span>
				</td></tr>
			</table>

			
				<table cellpadding="0" cellspacing="0" border="0" id="searchbox">
					<tr>
						<td nowrap="nowrap">
<%
                         if lcl_googleSearchID <> "" then
                           'Place this tag where you want the search box to render
                            'response.write "<p><gcse:searchbox-only></gcse:searchbox-only></p>" & vbcrlf
                            response.write "<gcse:search></gcse:search>" & vbrlf
                            response.write "<div style=""font-size: 1.5em; color: #ff0000;"">ONLY searches ""published documents""</div>" & vbcrlf
                         else
                            response.write "<form action=""search/search.asp"" method=""post"" name=""frmSearch"">" & vbcrlf
							response.write "<strong>Search:</strong>" & vbcrlf
							response.write "<input type=""hidden"" name=""Action"" value=""Go"" />" & vbcrlf
							response.write "<input type=""text"" id=""SearchString"" name=""SearchString"" size=""65"" maxlength=""100"" style=""background-color:#eeeeee;width:255px; height:19px; border:1px solid #000033;"" />" & vbcrlf
							response.write "<a href=""#"" onClick='ValidateSearch();'><img src=""../images/go.gif"" border=""0"" />" & langGo & "</a>" & vbcrlf
							response.write "</form>" & vbcrlf
                         end if
%>
						</td>
					</tr>
				</table>
			

			<div id="JQueryFTD_Demo" class="demo1">Loading, please wait.</div>

			<p><strong>Instructions: Left click the folders and files to open, close or view. Right click to perform actions. 
				</strong></p>

			<form name="frmAction" action="addfolder.asp" method="post">
				<input type="hidden" id="path" name="path" value="" />
			</form>

			<ul id="myFolderMenu" class="contextMenu">
<%				If session("orgregistration") And UserHasPermission( Session("UserId"), "public folder" ) Then %>
					<li class="editpublicaccess"><a href="#editpublicaccess">Edit Public Access</a></li>
<%				End If		
				If UserHasPermission( Session("UserId"), "internal folder" ) Then %>
					<li class="editsecurity"><a href="#editsecurity">Edit Security</a></li>
<%				End If		%>
				<li class="separator addfolder"><a href="#addfolder">Add Folder</a></li>
				<li class="adddocument"><a href="#adddocument">Add Document</a></li>
				<li class="separator deletefolder"><a href="#deletefolder">Delete Folder</a></li>
			</ul>

			<ul id="myFileMenu" class="contextMenu">
				<li class="rename"><a href="#rename">Rename</a></li>
				<li class="move"><a href="#move">Move</a></li>
				<li class="separator delete"><a href="#delete">Delete</a></li>
			</ul>

			<ul id="myRootMenu" class="contextMenu">
				<li class="separator addfoldertoroot"><a href="#addfoldertoroot">Add Folder</a></li>
			</ul>
		</div>
	</div>
	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  


</body>

</html>