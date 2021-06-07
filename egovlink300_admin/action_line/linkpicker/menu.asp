<%
'Get the city document location
 sLocationName = trim(GetVirtualName(Session("OrgID")))
%>
<html>
<head>
  <style type="text/css">
  <!--
    td      {font-family:MS Sans Serif,Tahoma,Arial; font-size:10px; color:#ffffff;}
    .sel    {border:1px solid #cccccc; cursor:hand; padding:5px 0px;}
    .notsel {border:1px solid #666666; cursor:hand; padding:5px 0px;}
  //-->
  </style>
  <script language="javascript">
  <!--
    function MakeActive(o) {
      document.all.exdoc.className  = "notsel";
      document.all.newdoc.className = "notsel";
      document.all.newurl.className = "notsel";
   	  document.all.newpay.className = "notsel";
      o.className = "sel";
      parent.MakeActive(o.id);
    }
  //-->
  </script>
</head>
<body bgcolor="#666666" topmargin="0" leftmargin="0">
<table border="0" cellpadding="0" cellspacing="0" width="85">
  <tr>
		    <td align="center">
       			<div class="sel" id="exdoc">
         			<img src="../../images/picker/existdoc.gif" /><br />
            Documents
          </div>
      </td>
  </tr>
</table>
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