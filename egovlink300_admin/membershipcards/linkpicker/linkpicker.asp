<%sLocationName = trim(GetVirtualName(session("orgid")))%>
<html>
<head>
  <title>Choose File...</title>
<style type="text/css">
  <!--
    td, input, select {font-family:MS Sans Serif,Tahoma,Arial; font-size:11px;}
  //-->
</style>
<script language="javascript">
  <!--
  function MakeActive(id) {
    document.all.exdoc.style.display  = "none";
    document.all.newdoc.style.display = "none";
    document.all.newurl.style.display = "none";
    document.all.newpay.style.display = "none";
    eval("document.all." + id + ".style.display = ''");
  }
    
  function saveSelection() {
    var objParent = window.opener;
    var path      = document.frmFilePath.FilePath.value;
 	  document.frmFilePath.FileLink.value = "<a href='" + document.all.currentpath.value + "/" + path + "'>" + path + "</a>";
	   objParent.addItem.itemID.value=document.all.currentpath.value + "/" + path;
	   objParent.addItem.link.value=path;
	   if(objParent.addItem.title.value=='') {
       objParent.addItem.title.value=path;
    }
    window.close();
  }

  function buildlink(sFormField) {
  		var sLink = "";
		  var path  = document.frmFilePath.FilePath.value;

  		if (document.frmFilePath.FilePath.value == "") {
     			alert("Please select a file.");
	     		return;
  		}
  		//sLink = "http://www.egovlink.com" + document.all.currentpath.value + "/" + path;
		  //sLink = sLink.replace("/custom/pub","");

    lcl_currentpath  = document.getElementById("currentpath");
    lcl_originalpath = document.getElementById("original_currentpath");

  		sLink = lcl_currentpath.value + "/" + path;
    sLink = sLink.replace(lcl_originalpath.value,"");

    eval("window.opener.document.getElementById('" + sFormField + "').value='" + sLink + "'");
	   window.close();
	 }

 	function buildactionlink(sFormField){
	   var sLocation = "<%=sLocationName%>";
	   var iFormID = document.frmAddArticle.iFormID.value;
	   var sLinkName = document.frmAddArticle.ALinkName.value;
	   if (sLinkName =='') {
      		sLinkName = document.frmAddArticle.AFormName.value;
    }
	   sLink = "<a target='_EGOVLINK' href='http://www.egovlink.com/" + sLocation + "/action.asp?actionid=" + iFormID + "'>" + sLinkName + "</a>";

 	  var oFormField = 'window.opener.document.' + sFormField; 
	   window.opener.insertAtCaret(eval(oFormField), sLink);
	   window.close();
 	}

	 function buildpaymentlink(sFormField){
	   var sLocation = "<%=sLocationName%>";
	   var iFormID = document.frmPaymentLink.iFormID.value;
 	  var sLinkName = document.frmPaymentLink.ALinkName.value;
	   if (sLinkName =='') {
		      sLinkName = document.frmPaymentLink.AFormName.value;
 	  }

	   sLink = "<a target='_EGOVLINK' href='http://www.egovlink.com/" + sLocation + "/payment.asp?paymenttype=" + iFormID + "'>" + sLinkName + "</a>";

	   var oFormField = 'window.opener.document.' + sFormField; 
	   window.opener.insertAtCaret(eval(oFormField), sLink);
 	  window.close();
	 }

 	function buildwebpagelink(sFormField){
	   var sURL = document.frmURL.UrlType.value + document.frmURL.Url.value;
 	  var sLinkName = document.frmURL.UrlName.value;
	   if (sLinkName =='') {
		      sLinkName = document.frmURL.Url.value;
 	  }
	   sLink = "<a target='_EGOVLINK' href='" + sURL + "'>" + sLinkName + "</a>";

 	  var oFormField = 'window.opener.document.' + sFormField; 
	   window.opener.insertAtCaret(eval(oFormField), sLink);
	   window.close();
 	}

  function myFunction() {
	   alert(document.frmAddArticle.currentfolderpath.value);
	 }

  function setupOriginalCurrentPath() {
    lcl_currentpath  = document.getElementById("currentpath");
    lcl_originalpath = document.getElementById("original_currentpath");

    if(lcl_originalpath.value == "") {
       lcl_originalpath.value = lcl_currentpath.value;
    }
  }
//-->
</script>
</head>
<body bgcolor="#d4d0c8" leftmargin="2" topmargin="0" onload="setupOriginalCurrentPath()">
  <input type="hidden" name="currentpath" id="currentpath" size="80" /><br />
  <input type="hidden" name="original_currentpath" id="original_currentpath" size="80" />
<table border="0" cellpadding="3" cellspacing="0">
  <tr>
      <td>&nbsp;</td>
      <td>
          <input type="text" name="currentfolder" style="height:20px; width:250px;" readonly />&nbsp;
          <!-- <a href="#" style="color:#0000ff" onclick="explorer.window.history.back();" name="anchorBack"><img src="images/up.gif" alt="Back" border="0" align="absmiddle" /></a> -->
          <img src="../../images/picker/up.gif" alt="Back" border="0" align="absmiddle" style="cursor:pointer" onclick="explorer.window.history.back();" />
      </td>
  </tr>
  <tr>
      <td rowspan="3" valign="top">
          <iframe name="menu" width="100" height="265" src="menu.asp"></iframe>
      </td>
  </tr>
  <tr>
      <td valign="top">
          <iframe id="explorer" name="explorer" width="400" height="250" src="loadtree.asp?path=<%="/public_documents300/custom/pub/" & sLocationName & "/unpublished_documents" %>"></iframe>

          <div id="exdoc">
          <table border="0" cellpadding="0" cellspacing="0" width="400">
            <form name="frmFilePath" id="frmFilePath">
            <tr>
                <td valign="top" style="padding-top:5px;">
               			  <table>
                				  <tr>
                          <td>File name:&nbsp;&nbsp;</td>
                          <td><input type="text" name="FilePath" id="FilePath" style="width:250px; height:20px;" readonly /></td>
                      </tr>
                    </table>
                    <br />
                  <%
                    if request("message") <> "" then
                       response.write request("message") & vbcrlf
                    end if
                  %>
                </td>
                <td valign="top" align="right" style="padding-top:5px;">
                    <% displayButtons "exdoc",request("fid") %>
                </td>
            </tr>
            </form>
          </table>
          </div>

          <div id="newdoc" style="display:none">
          <table border="0" cellpadding="0" cellspacing="0" width="400">
            <form name="frmAddArticle">
  		          <input type="hidden" name="currentfolderpath" />
         		   <input type="hidden" name="iFormID" />
            <tr>
                <td valign="top" style="padding-top:5px;">
               			  <table>
                				  <tr>
                    				  <td>Link Text:&nbsp;&nbsp;</td>
                    				  <td><input type="text" name="ALinkName" style="width:250px; height:20px;" /></td>
                    		</tr>
                				  <tr>
                    				  <td>Form Name:&nbsp;&nbsp;</td>
                    				  <td><input type="text" name="AFormName" style="width:250px; height:20px;" readonly /></td>
                    		</tr>
              			   </table>
                  	 <br />
                  <%
                    if request("message") <> "" then
                       response.write request("message") & vbcrlf
                    end if
                  %>
                </td>
                <td valign="top" align="right" style="padding-top:5px;">
                    <% displayButtons "newdoc",request("fid") %>
                </td>
            </tr>
            </form>
          </table>
          </div>

          <div id="newpay" style="display:none">
          <table border="0" cellpadding="0" cellspacing="0" width="400">
            <form name="frmPaymentLink">
         		   <input type="hidden" name="currentfolderpath" />
		            <input type="hidden" name="iFormID" />
            <tr>
                <td valign="top" style="padding-top:5px;">
             	  		  <table>
              		  		  <tr>
                    				  <td>Link Text:&nbsp;&nbsp;</td>
                  		  		  <td><input type="text" name="ALinkName" style="width:250px; height:20px;" /></td>
                    		</tr>
                				  <tr>
                    				  <td>Form Name:&nbsp;&nbsp;</td>
                    				  <td><input type="text" name="AFormName" style="width:250px; height:20px;" readonly /></td>
                    		</tr>
               			  </table>
                  		<br />
                  <%
                    if request("message") <> "" then
                       response.write request("message") & vbcrlf
                    end if
                  %>
                </td>
                <td valign="top" align="right" style="padding-top:5px;">
                    <% displayButtons "newpay",request("fid") %>
                </td>
            </tr>
            </form>
          </table>
          </div>

          <div id="newurl" style="display:none">
      		  <table border="0" cellpadding="0" cellspacing="0" width="400">
            <form name="frmURL">
            <tr>
                <td valign="top" style="padding-top:5px;">
                    Name:&nbsp;&nbsp;<input type="text" name="UrlName" style="width:267px; height:20px;" />
                    <br />
                    <img src="../../images/picker/spacer.gif" width="1" height="5" />
                    <br />
                    URL:&nbsp;&nbsp;&nbsp;&nbsp;
                    <select name="UrlType">
                      <option value="http://">http://</option>
                      <option value="mailto:">mailto:</option>
                      <option value="ftp://">ftp://</option>
                    </select>
                    <input type="text" name="Url" style="width:207px; height:20px;" />
                </td>
                <td align="right" style="padding-top:5px;">
                    <% displayButtons "newurl",request("fid") %>
                </td>
            </tr>
        		  </form>
          </table>
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
  
  sSQL = "SELECT OrgVirtualSiteName FROM Organizations WHERE orgid = " & iorgid
  set oRst = Server.CreateObject("ADODB.Recordset")
  oRst.open sSQL,Application("DSN"),3,1
  
  if not oRst.eof then
    	sReturnValue = oRst("OrgVirtualSiteName")
  end if

  GetVirtualName = sReturnValue

end function

'------------------------------------------------------------------------------
sub displayButtons(iDivID,iFieldID)
  lcl_button_label = ""
  lcl_onclick      = ""

  if iDivID <> "" AND iFieldID <> "" then
     if UCASE(iDivID) = "EXDOC" then
        lcl_button_label = "Select"
        lcl_onclick      = "buildlink"
     elseif UCASE(iDivID) = "NEWDOC" then
        lcl_button_label = "Add Link"
        lcl_onclick      = "buildactionlink"
     elseif UCASE(iDivID) = "NEWPAY" then
        lcl_button_label = "Add Link"
        lcl_onclick      = "buildpaymentlink"
     elseif UCASE(iDivID) = "NEWURL" then
        lcl_button_label = "Add Link"
        lcl_onclick      = "buildwebpagelink"
     end if

     response.write "<input type=""button"" value=""" & lcl_button_label & """ style=""width:80px; height:22px;"" onClick=""javascript:" & lcl_onclick & "('" & iFieldID & "');"" /><br />" & vbcrlf
     response.write "<img src=""../images/picker/spacer.gif"" width=""1"" height=""5"" /><br />" & vbcrlf
     response.write "<input type=""button"" value=""Cancel"" style=""width:80px; height:22px;"" onclick=""window.close();"" />" & vbcrlf

  end if

end sub
%>