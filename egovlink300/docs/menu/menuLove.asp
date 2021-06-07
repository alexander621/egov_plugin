<!-- #include file="loadfolder_db.inc" //-->

<html>
<head>
  <title>E-Gov Services - Loveland, Ohio</title>
  <link rel="stylesheet" href="../../css/styles.css" type="text/css">
  <script language="Javascript" src="menu.js" type="text/javascript"></script>
  <script language="Javascript">
  <!--
    var pageLoad = false;

    function doError() {
      if (!pageLoad) {
        pageLoad = true;
      }
      else {
        obj = hiddenframe.chunk;
        if (obj == undefined) {
          loadFrame(true);
          s = eCurrentLI.id;
          plen = <%= Len(Application("eCapture_ArticlesPath")) + 1 %>;
          pos = s.indexOf("<%= Application("eCapture_ArticlesPath") & "/" %>");
          if (pos >= 0)
            s = s.substr(0,pos) + s.substr(pos+plen,s.length-pos+plen);
          parent.fraTopic.document.location.href = "../error.asp?path=" + s;
        }
      }
    }
  //-->
  </script>
  <link href="menu.css" type="text/css" rel="stylesheet">
  
</head>

<!--#Include file="../../include_top.asp"-->

<!--BODY CONTENT-->

<TR><TD VALIGN=TOP>
  <p class=title>Online Documents</p>

  <div id="menu2" style="margin-left:20px;" >
    <table border="0" cellpadding="0" cellspacing="0" width="100%" id="menu">
      <tr>
        <td valign="top" width="100%">
          <nobr>
            <ul id="ulRoot" style="display: block;">
			        <%=LoadFolder( Application("eCapture_ArticlesPath") )%>
            </ul>
          </nobr>
        </td>
      </tr>
    </table>
    <iframe name="hiddenframe" src="" width="0" height="0" onload="doError();"></iframe>
  </div>
  

  
