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
    
    function saveSelection()
    {
      var objParent=window.opener;
      var path=document.frmFilePath.FilePath.value;
	  objParent.addItem.itemID.value=document.all.currentpath.value + "/" + path;
	  objParent.addItem.link.value=path;
	  if(objParent.addItem.title.value=='')objParent.addItem.title.value=path;
      window.close();
    }
    function myFunction() {
	    alert(document.frmAddArticle.currentfolderpath.value);
	 }
  //-->
  </script>
</head>

<body bgcolor="#d4d0c8" leftmargin="2" topmargin="0">
<input type="hidden" name="currentpath" >
  <table border="0" cellpadding="3" cellspacing="0">
    <tr>
      <td>&nbsp;</td>
      <td>
        <input type="text" name="currentfolder" style="height:20px; width:250px;" readonly>&nbsp;<a href="#" style="color:#0000ff" onclick="explorer.window.history.back();" name="anchorBack"><img src="images/up.gif" alt="Back" border="0" align="absmiddle"></a>
      </td>
    </tr>
    <tr>
      <td rowspan="3" valign="top">
        <iframe name="menu" width="90" height="250" src="menu.asp"></iframe>
      </td>
    </tr>
    <tr>
      <td valign="top">
        <iframe name="explorer" width="400" height="250" src="loadtree.asp?path=<%=Application("eCapture_ArticlesPath") %>"></iframe>

        <div id="exdoc">
          <table border="0" cellpadding="0" cellspacing="0" width="400">
          <form name="frmFilePath">
            <tr>
              <td valign="top" style="padding-top:5px;">
                File name:&nbsp;&nbsp;<input type="text" name="FilePath" style="width:250px; height:20px;" readonly>
		<br>
			<% if request("Message") <> "" then %>
				<% =request("Message") %>
			<% end if %>
              </td>
              <td align="right" style="padding-top:5px;">
                <input type="button" value="Attach" style="width:80px; height:22px;" onClick="javascript:saveSelection();"><br>
                <img src="images/spacer.gif" width="1" height="5"><br>
                <input type="button" value="Cancel" style="width:80px; height:22px;" onclick="window.close();">
              </td>
            </tr>
            </form>
          </table>
        </div>

        <div id="newdoc" style="display:none">
          <table border="0" cellpadding="0" cellspacing="0" width="400">
          <form name="frmAddArticle" action="upload.asp?task=UPLOAD" method="POST" enctype="multipart/form-data">
            <input type="hidden" name="currentfolderpath">
            <tr>
              <td valign="top" style="padding-top:5px;">
                File name:&nbsp;&nbsp;<input type="file" name="binFile" style="width:250px; height:20px;">
		<br>
			<!--a onclick="javascript:myFunction();">TEST</a-->
			<% if request("Message") <> "" then %>
				<% =request("Message") %>
			<% end if %>
              </td>
              <td align="right" style="padding-top:5px;">
                <input type="submit" value="Attach" style="width:80px; height:22px;"><br>
                <img src="images/spacer.gif" width="1" height="5"><br>
                <input type="button" value="Cancel" style="width:80px; height:22px;" onclick="window.close();">
              </td>
            </tr>
	    </form>
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
        </div>


      </td>
    </tr>
  </table>
</body>
</html>
