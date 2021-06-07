<%Response.buffer = true%>
<!-- #include file="URLDecode.asp" //-->
<!-- #include file="loadfolder.inc" //-->
<!-- #include file="../../includes/common.asp" //-->
<%

strPath = URLDecode( request("path"))
Session("curpath") = strPath


' GET CITY DOCUMENT LOCATION
sLocationName = trim(GetVirtualName(Session("OrgID")))
		   
' DISPLAY FOLDER TREE %>
<% strList = LoadFolder("/public_documents300/custom/pub/" & sLocationName  , strPath  )



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
<!-- <body leftmargin="0" topmargin="0" onload="parent.document.frmAddArticle.currentfolderpath.value='<%=strPath%>';parent.document.all.currentfolder.value='<%=sFolder%>';parent.document.all.currentpath.value='<%=strPath%>'"> -->
<body leftmargin="0" topmargin="0" onload="parent.document.all.currentfolderpath.value='<%=strPath%>';parent.document.all.currentfolder.value='<%=sFolder%>';parent.document.all.currentpath.value='<%=strPath%>'">
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
              End If
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
Function GetVirtualName(iorgid)
  
  sReturnValue = "UNKNOWN"
  
  Set oRst = Server.CreateObject("ADODB.Recordset")
  sSQL = "SELECT OrgVirtualSiteName FROM Organizations WHere orgid = " &  iorgid 
  oRst.open sSQL,Application("DSN"),3,1
  
  If NOT oRst.EOF THEN
	sReturnValue = oRst("OrgVirtualSiteName")
  END IF

  GetVirtualName = sReturnValue

End Function
%>