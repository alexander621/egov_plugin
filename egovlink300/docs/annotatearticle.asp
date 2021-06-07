<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'if not logged in dont allow access, no matter what
If Session("UserID") = 0 Then Response.Redirect RootPath

 sSql = "EXEC ListAnnotations " & session("OrgID") &", '" & request("path") & "'"
 Set oRst = Server.CreateObject("ADODB.Recordset")
 oRst.Open sSql, Application("DSN"), 3, 1

   If Not oRst.EOF Then
	  intDocumentID = oRst("DocumentID")
	  Do While Not oRst.EOF
	   If sBgcolor = "#ffffff" Then sBgcolor = "#eeeeee" Else sBgcolor = "#ffffff"
		strText = strText &  "<tr bgcolor=""" & sBgcolor & """><td>&nbsp;</td><td valign=top ><input type=""checkbox"" class=""nomargin"" name=""del_" & oRst("AnnotationID") & """></td><td nowrap valign=top>"& oRst("FirstName")& " " &oRst("LastName") & "</td><td valign=top>"&oRst("AnnotationText")&"</td><td valign=top nowrap>"&oRst("AnnotationDateTime")&"</td></tr>"
	  oRst.MoveNext
	  Loop
      oRst.Close
	  Set oRst = Nothing
   Else
        ' Document for create new annotations
		sSql = "EXEC GetDocumentIDbyPath " & session("OrgID") & ", '" & request("path") & "'"
		Set oRst = Server.CreateObject("ADODB.Recordset")
		oRst.Open sSql, Application("DSN"), 3, 1
		intDocumentID = oRst("DocumentID")
		strText="<tr><td colspan=4><font class=empty>No annotations for this document.</font></td></tr>"
       oRst.Close
	   Set oRst = Nothing
   End If

%>

<html>
<head>
  <title>Annotations</title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script src="../scripts/selectAll.js"></script>
  <style type="text/css">
  <!--
    .nomargin {margin:-4px;}
  //-->
  </style>
  <SCRIPT LANGUAGE="JavaScript">
		<!--// This function will open page in new window
		function NewWindow(page) {
			OpenWin = this.open(page, "CtrlWindow", "height=385,width=400,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no");
		}
		// End -->
  </SCRIPT>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <td><font size="+1"><b><%=langTabDocuments%>: Annotations</b></font><br>
      There are _ annotations for document <b>document name</b>.
	  </td>
    </tr>
    <tr>
     <td valign="top">
        <form name="frmAnnotationList" action="delannotations.asp" method="post">
          <div style="font-size:10px; padding-bottom:5px;"><img src="../images/arrow_back.gif" align="absmiddle">
          
          <%
          If iPage > 1 Then
            Response.Write "<a href=""default.asp?page=" & iPage-1 & """>" & langPrev & "  " & Session("PageSize") & "</a>&nbsp;&nbsp;"
          Else
            Response.Write "<font color=""#999999"">Prev " & Session("PageSize") & "</font>&nbsp;&nbsp;"
          End If

          If iNumMoreRecords > 0 Then
            Response.Write "<a href=""default.asp?page=" & iPage+1 & """>" & langNext & " " & Session("PageSize") & "</a>"
          Else
            Response.Write "<font color=""#999999"">Next " & Session("PageSize") & "</font>"
          End If
          %>
		  <img src="../images/arrow_forward.gif" align="absmiddle">
		  &nbsp;&nbsp;&nbsp;&nbsp;
		  
		  <!--Begin New Button -->
		  <% If HasPermission("CanEditAnnotation") Then %>
		  <img class="menuimage" src="menu/images/annotate.gif" width="18" height="18" align="absmiddle">
		  <a href="javascript:NewWindow('newAnnotation.asp?ID=<%=intDocumentID%>');">New Annotation</a>
		  <% End If %>
		  <!--End New Button -->
         
		  <!--Begin Delete Button-->
		  <% If HasPermission("CanEditAnnotation") Then %>
          &nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/small_delete.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.frmAnnotationList.submit();">Delete</a>
		  <% End If %>
		  <!--End Delete Button-->
          
		  </div>

          <table width="100%" cellpadding="5" cellspacing="0" border="0" class="tablelist">
            <tr>
              <th width="1%">&nbsp;</th>
			  <th align="center" >&nbsp;</th>
              <th align="left">User</th>
              <th align="left" width="70%">Annotation</th>
              <th align="left"><%=langDate%></th>
            </tr>
            <%= strText %>
          </table>
        </form>
      </td>
    </tr>
  </table>

</body>
</html>

