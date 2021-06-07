<%
' GET CITY DOCUMENT LOCATION
sLocationName = trim(GetVirtualName(Session("OrgID")))
%>

<html>
<head>
  <style type="text/css">
  <!--
    td {font-family:MS Sans Serif,Tahoma,Arial; font-size:10px; color:#ffffff;}
    .sel {border:1px solid #cccccc; cursor:hand; padding:5px 0px;}
    .notsel {border:1px solid #666666; cursor:hand; padding:5px 0px;}
  //-->
  </style>
  <script language="Javascript">
  <!--
    function MakeActive(o) {
      document.all.exdoc.className = "notsel";
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
    <tr><td align="center"><div class="sel" id="exdoc" onclick="MakeActive(this);parent.explorer.window.location.href='loadtree.asp?path=<%="/public_documents300/custom/pub/" & sLocationName & "/published_documents" %>';" ><img src="images/existdoc.gif"><br>Link Document</div></td>
    </tr>
  
   <tr>
      <td align="center"><div class="notsel" id="newdoc" onclick="MakeActive(this); parent.explorer.window.location.href='listforms.asp';"><img src="images/newdoc.gif"><br>Link Action Form</div></td>
    </tr>
	<tr>
      <td align="center"><div class="notsel" id="newpay" onclick="MakeActive(this); parent.explorer.window.location.href='listpayments.asp';"><img src="images/newdoc.gif"><br>Link Payment Form</div></td>
    </tr>
    <tr>
      <td align="center"><div class="notsel" id="newurl" onclick="MakeActive(this);"><img src="images/newurl.gif"><br>Link to URL<br><br></div></td>
    </tr>


  </table>
</body>
</html>

<%
Function GetVirtualName(iorgid)
  
  sReturnValue = "UNKNOWN"
  
  Set oRst = Server.CreateObject("ADODB.Recordset")
  sSQL = "SELECT OrgVirtualSiteName FROM Organizations WHere orgid='" &  iorgid & "'"
  oRst.open sSQL,Application("DSN"),3,1
  
  If NOT oRst.EOF THEN
	sReturnValue = oRst("OrgVirtualSiteName")
  END IF

  GetVirtualName = sReturnValue

End Function
%>