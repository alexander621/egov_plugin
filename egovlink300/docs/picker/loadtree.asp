<%Response.buffer = true%>
<!-- #include file="URLDecode.asp" //-->
<!-- #include file="loadfolder_DB.inc" //-->
<%
Dim strPath, strList, sFolder
strPath = URLDecode( request("path"))
Session("curpath") = strPath
strList = LoadFolder( strPath )

sFolder = Right(strPath, Len(strPath) - InStrRev(strPath,"/"))
If sFolder = "pub" Then
  sFolder = "Documents"
End If
If Left(sFolder,2) = "z." Then
  sFolder = Mid(sFolder, 3)
End If
%>

<html>
<head>
  <link href="menu.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" onload="parent.document.all.currentFolderPath.value='<%=strPath%>';parent.document.all.currentfolder.value='<%=sFolder%>'">
  <div id="menu2">
    <table border="0" cellpadding="0" cellspacing="0" width="100%" id="menu">
      <tr>
        <td valign="top" width="100%">
          <nobr>
            <ul id="ulRoot" style="display: block;">
              <%
              If strList = "" Then
                Response.Write "&nbsp;<font color=""#003366""><i>Permission Denied</i></font>"
              Else
                Response.Write strList
				%>
			    <div id="newfolder" style="display:none">
				<br>
				<form name="frmNewFolder" action="../newfolder.asp" method="post" target="_top"> 
	             <img src="../../images/picker/newfolder.gif"><input type=text name=folderName value='new folder' style="border: none" onBlur="frmNewFolder.submit();">
				 </form>
				</div>
              <%End If
              %>
            </ul>
          </nobr>
        </td>
      </tr>
    </table>
    <iframe name="hiddenframe" src="" width="0" height="0"></iframe>
  </div>
  
	
</body>
</html>