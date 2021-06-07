<%Response.buffer = true%>
<!-- #include file="URLDecode.asp" //-->
<!-- #include file="linkpicker_global_functions.asp" //-->
<!-- #include file="../../includes/common.asp" //-->
<%
strPath = URLDecode( request("path"))
session("curpath") = strPath

'Get city document location
 sLocationName = trim(GetVirtualName(session("orgid")))
		   
'Display Folder Tree
 strList = LoadFolder("/public_documents300/custom/pub/" & sLocationName, strPath  )
 sFolder = Right(strPath, Len(strPath) - InStrRev(strPath,"/"))

 if sFolder = "pub" then
    sFolder = "Documents"
 end if

 if Left(sFolder,2) = "z." then
    sFolder = Mid(sFolder, 3)
 end if
%>
<html>
<head>
  <link href="menu.css" rel="stylesheet" type="text/css">
</head>
<!-- <body leftmargin="0" topmargin="0" onload="parent.document.frmAddArticle.currentfolderpath.value='<%=strPath%>';parent.document.all.currentfolder.value='<%=sFolder%>';parent.document.all.currentpath.value='<%=strPath%>'"> -->
<body leftmargin="0" topmargin="0" onload="parent.document.all.currentfolderpath.value='<%=strPath%>';parent.document.all.currentfolder.value='<%=sFolder%>';parent.document.all.currentpath.value='<%=strPath%>'">

<div id="menu2">
<table border="0" cellpadding="0" cellspacing="0" width="100%" id="menu">
  <tr>
      <td valign="top" width="100%">
          <nobr>
          <ul id="ulRoot" style="display: block;">
          <%
            if strList = "" then
               response.write "&nbsp;<font color=""#003366""><i>Permission Denied</i></font>"
            else
               response.write strList
            end if
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
<%
'------------------------------------------------------------------------------
function GetVirtualName(iorgid)

  sReturnValue = "UNKNOWN"

  sSQL = "SELECT OrgVirtualSiteName FROM Organizations WHERE orgid = " & clng(iorgid)
  set oRst = Server.CreateObject("ADODB.Recordset")
  oRst.open sSQL,Application("DSN"),3,1

  if not oRst.eof then
    	sReturnValue = oRst("OrgVirtualSiteName")
  end if

  GetVirtualName = sReturnValue

end function
%>