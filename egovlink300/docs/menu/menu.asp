<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="loadfolder_db.inc" //-->

<html>
<head>
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

<body text="#ffffff" bgcolor="#ffffff" leftmargin="4" topmargin="0" marginwidth="0" marginheight="0">
  <br>
  <div id="menu2">
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
  
  ' Set context sensitive menu options
   <div id="mnu_Category" class="skin0" onMouseover="highlightie5(event)" onMouseout="lowlightie5(event)" onClick="jumptoie5(event)" style="visibility:hidden;display:none">
     <% If HasPermission("CanEditDocuments") Then %>
    <div class="menuitems" url="../members.asp"><img class="menuimage" src="../../images/newpermission.gif" width="18" height="18" align="absmiddle"> Edit Security</div>
    <hr size="1" color="#999999">
    <div class="menuitems" url="../addfolder.asp"><img class="menuimage" src="images/folder_closed.gif" width="18" height="18" align="absmiddle"> Add Folder</div>
    <div class="menuitems" url="../addarticle.asp"><img class="menuimage" src="images/document.gif" width="18" height="18" align="absmiddle"> Add Document</div>
    <div class="menuitems" url="../addhelp.asp"><img class="menuimage" src="images/helpdocument.gif" width="18" height="18" align="absmiddle"> Add Help</div>
    <hr size="1" color="#999999">
    <div class="menuitems" url="../delcategory.asp"><img class="menuimage" src="images/delete.gif" width="18" height="18" align="absmiddle"> Delete Folder</div>
    <%Else%>
		<div class="menuitems" url="../main.asp" target="_top" ><img class="menuimage" src="images/delete.gif" width="18" height="18" align="absmiddle"></div>		
    <%End If%>
   </div>

  <div id="mnu_Article" class="skin0" onMouseover="highlightie5(event)" onMouseout="lowlightie5(event)" onClick="jumptoie5(event)" style="visibility:hidden;display:none">
     <% If HasPermission("CanEditDocuments") Then %>
    <div class="menuitems" url="../delarticle.asp"><img class="menuimage" src="images/delete.gif" width="18" height="18" align="absmiddle"> Delete Document</div>
	
	<%If HasPermission("CanViewAnnotation") Then%>
	<div class="menuitems" url="../annotatearticle.asp"><img class="menuimage" src="images/annotate.gif" width="18" height="18" align="absmiddle"> Annotations </div>
	<%End If%>
	
	<%ELSE%>
    <div class="menuitems" url="../main.asp" target="_top" ><img class="menuimage" src="images/delete.gif" width="18" height="18" align="absmiddle"></div>
    <% End If%>
  </div>
  
  <%' if user has no permissions do not display context sensitive menu
  If HasPermission("CanEditDocuments") Then %>
	<script>initContextMenu();</script>
  <% End If %>
 
  
</body>
</html>
