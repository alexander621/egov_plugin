<html>
<head>
  <title>Loading...</title>
  <link href="global.css" rel="stylesheet">
</head>

<body onload="document.location.href='<%= Request("file") %>'">
  <table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
    <tr>
      <td align="center" valign="center">
        <font face="Verdana" size="1">
          <% If Request("prompt") <> "" Then Response.Write Request("prompt") Else Response.Write "Loading" %>...<br>
        <br>
        <img src="../images/progressbar.gif" width="130" height="12" border="0"><br>
        <% If Request("wait") = 1 Then Response.Write "(This may take a few minutes)" %>
        </font>
      </td>
    </tr>
  </table>
</body>
</html>