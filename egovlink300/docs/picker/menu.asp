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
      o.className = "sel";
      parent.MakeActive(o.id);
    }
  //-->
  </script>
</head>

<body bgcolor="#666666" topmargin="0" leftmargin="0">
  <table border="0" cellpadding="0" cellspacing="0" width="85">
    <tr>
      <td align="center"><div class="sel" id="exdoc" onclick="MakeActive(this);"style="height:75px; vertical-align:middle;"><img src="images/existdoc.gif"><br>Documents</div></td>
    </tr>
    <!--<tr>
      <td align="center"><div class="notsel" id="newdoc" onclick="MakeActive(this);"><img src="images/newdoc.gif"><br>Upload Document</div></td>
    </tr>
    <tr>
      <td align="center"><div class="notsel" id="newurl" onclick="MakeActive(this);"><img src="images/newurl.gif"><br>URL<br><br></div></td>
    </tr>//-->
  </table>
</body>
</html>