<%
sType = Request.QueryString("type")

If request.querystring("task")<>"return" Then
	strFolderList = Application("eCapture_ArticlesPath")
Else
	strFolderList = request.querystring("path")
End If
%>

<html>
<head>
  <title>Choose File...</title>
  <style type="text/css">
  <!--
    td, input, select {font-family:MS Sans Serif,Tahoma,Arial; font-size:11px;}
  //-->
  </style>
  <script language="Javascript">
  <!--
    function MakeActive(id) {
      document.all.exdoc.style.display = "none";
      document.all.newdoc.style.display = "none";
      document.all.newurl.style.display = "none";
  	  eval("document.all." + id + ".style.display = ''");
    }

    function NewFolder() {
      document.all.explorer.contentWindow.document.all.newfolder.style.display = '';
      document.all.explorer.contentWindow.document.all.frmNewFolder.folderName.select();
    }

    function doFolderSelect() {
      eval("window.opener.txtDir.value='" + document.all.currentFolderPath.value + "/')");
      eval("window.opener.DrawTree('" + document.all.currentFolderPath.value + "/')");
      window.close();
    }
  //-->
  </script>
</head>

<body bgcolor="#d4d0c8" leftmargin="2" topmargin="0">
<form name="test">
  <input type="hidden" name="currentFolderPath">
  <table border="0" cellpadding="3" cellspacing="0">
    <tr>
      <td>&nbsp;</td>
      <td>
        <input type="text" name="currentfolder" style="height:20px; width:250px;" readonly>&nbsp;<a href="#" style="color:#0000ff" onclick="explorer.window.history.back();" name="anchorBack"><img src="images/up.gif" alt="Back" border="0" align="absmiddle"></a>
		    <!--<a href="#" onclick="NewFolder(); return false;"><img src="../../images/picker/newfolder.gif" align="absmiddle" border=0 alt="New Folder"></a>//-->
	
      </td>
    </tr>
    <tr>
      <td rowspan="3" valign="top">
        <iframe name="menu" width="90" height="250" src="menu.asp"></iframe>
      </td>
    </tr>
    <tr>
      <td valign="top">
        <iframe name="explorer" width="400" height="250" src="loadtree.asp?path=<%=strFolderList%>">
		</iframe>

        <div id="exdoc">
          <table border="0" cellpadding="0" cellspacing="0" width="400">
            <tr>
              <td valign="top" style="padding-top:5px;">
               <% If sType <> "folder" Then %>
                File name:&nbsp;&nbsp;<input type="text" name="FilePath" style="width:250px; height:20px;" readonly>
              <% End If %>
              </td>
              <td align="right" style="padding-top:5px;">
               <% If sType = "folder" Then %>
                <input type="button" value="Select" style="width:80px; height:22px;" onclick="doFolderSelect();"><br>
              <% Else %>
                <input type="submit" value="Attach" style="width:80px; height:22px;"><br>
              <% End If %>
                <img src="images/spacer.gif" width="1" height="5"><br>
                <input type="button" value="Cancel" style="width:80px; height:22px;" onclick="window.close();">
              </td>
            </tr>
          </table>
        </div>
<!--
        <div id="newdoc" style="display:none">
          <table border="0" cellpadding="0" cellspacing="0" width="400">
            <tr>
              <td valign="top" style="padding-top:5px;">
                File name:&nbsp;&nbsp;<input type="file" name="Name" style="width:250px; height:20px;">
              </td>
              <td align="right" style="padding-top:5px;">
                <input type="submit" value="Attach" style="width:80px; height:22px;"><br>
                <img src="images/spacer.gif" width="1" height="5"><br>
                <input type="button" value="Cancel" style="width:80px; height:22px;" onclick="window.close();">
              </td>
            </tr>
          </table>
        </div>

        <div id="newurl" style="display:none">
          <table border="0" cellpadding="0" cellspacing="0" width="400">
            <tr>
              <td valign="top" style="padding-top:5px;">
                Name:&nbsp;&nbsp;<input type="text" name="UrlName" style="width:267px; height:20px;"><br>
                <img src="images/spacer.gif" width="1" height="5"><br>
                URL:&nbsp;&nbsp;&nbsp;&nbsp;<select name="UrlType">
                  <option value="http://">http://</option>
                  <option value="mailto:">mailto:</option>
                  <option value="ftp://">ftp://</option>
                </select>
                <input type="text" name="Url" style="width:207px; height:20px;">
              </td>
              <td align="right" style="padding-top:5px;">
                <input type="submit" value="Attach" style="width:80px; height:22px;"><br>
                <img src="images/spacer.gif" width="1" height="5"><br>
                <input type="button" value="Cancel" style="width:80px; height:22px;" onclick="window.close();">
              </td>
            </tr>
          </table>
        </div>//-->
         
        
		<div id="newfold" style="display:none">
   		  <table border="0" cellpadding="0" cellspacing="0" width="400">
            <tr>
              <td valign="top" style="padding-top:5px;">
			     <form name="frmNewFolder" action="../newfolder.asp" method="post"> 
                Folder name:&nbsp;&nbsp;<input type="text" name="FolderName" style="width:250px; height:20px;">
              </td>
              <td align="right" style="padding-top:5px;">
			    <input type="submit" value="Create" style="width:80px; height:22px;"><br>
                <img src="images/spacer.gif" width="1" height="5"><br>
                <input type="button" value="Cancel" style="width:80px; height:22px;" onclick="window.close();">
              </td>
            </tr>
          </table>
		</div>
		
 

      </td>
    </tr>
  </table>
  </form>
</body>
</html>