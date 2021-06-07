<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<!-- #include file="includes/time.asp" //-->
<!-- #include file="include_top_functions.asp" //-->
<!-- #include file="class/classOrganization.asp" -->
<!-- #include file="postings_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: postings_submit_userbids.asp
' AUTHOR:   David Boyer
' CREATED:  10/16/08
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This screen allows the user, if logged in, to upload their bid(s) to a posting
'
' MODIFICATION HISTORY
' 1.0  10/16/08	 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("job_postings,bid_postings") = "Y" then
    response.redirect "outage_feature_offline.asp"
 end if

 dim lcl_upload_filename

'Determine if the org requires the user to be logged in.
'If "yes", check to see if the user has logged in.
'If "yes", check to see if the user logged in has subscribed/registered to the category of the posting.
'If "yes" then show the "Submit Bid" button.
 if lcl_require_login then
    if request.cookies("userid") <> "" then
       if isCategoryAssigned(request.cookies("userid"),lcl_dlistid) then
          lcl_canUpload = "Y"
       end if
    end if
 else
    lcl_canUpload = "Y"
 end if

'option explicit 
 Response.Expires     = -1
 Server.ScriptTimeout = 600

 lcl_hidden = "text"  'Show/Hide all hidden fields.  TEXT=Show, HIDDEN=Hide

'Retrieve the org_group_id of the organization group that is to be maintained.
'If no value exists then redirect them back to the main results screen
 if request("posting_id") <> "" AND isnumeric(request("posting_id")) then
    lcl_posting_id = CLng(request("posting_id"))
 else
    response.redirect "postings.asp?listtype=" & request("listtype")
 end if

 if request("dlistid") <> "" then
    lcl_dlistid = request("dlistid")
 else
    lcl_dlistid = ""
 end if

 lcl_list_type = request("listtype")

 if lcl_list_type = "JOB" then
    lcl_list_label = "Job"
    lcl_list_title = "Job Postings"
 elseif lcl_list_type = "BID" then
    lcl_list_label = "Bid"
    lcl_list_title = "Bid Postings"
 end if

 lcl_user_label = request("userlabel")
 lcl_totalbytes = request.TotalBytes

 'if lcl_totalbytes <= 10485760 then  'APPX. 10MB
 'if lcl_totalbytes > 10240 then  'APPX. 10MB
 '   response.redirect sEgovWebsiteURL & "/postings_submit_userbids.asp?posting_id=" & lcl_posting_id & "&listtype=" & lcl_list_type & "&dlistid=" & lcl_dlistid & "&success=BIG&filesize=" & lcl_totalbytes
 'end if

'Determine if org requires user to be logged in to view the following field(s):
'  a. Download Available
' Track_DBsafe put in to block SQL Injection vulnerability SJL 1/2/2013
 lcl_require_login = orghasfeature(iorgid,Track_DBsafe(lcase(lcl_list_type))&"postings_requirepubliclogin")

	Dim oOrg
	set oOrg = New classOrganization

' Track_DBsafe put in to block SQL Injection vulnerability SJL 1/2/2013
 lcl_feature_name = oOrg.GetOrgFeatureName(Track_DBsafe(lcase(lcl_list_type))&"_postings")

'------------------------------------------------------------------------------
'Display all of the files in the directory *** FOR DEBUGGING ONLY ***
 'set MyFile   = server.createobject("scripting.filesystemobject")
 'set MyFolder = MyFile.GetFolder("d:\wwwroot\www.cityegov.com\egovlink300_admin\")
 'set MyFolder = MyFile.GetFolder("d:\wwwroot\www.cityegov.com\egovlink release QA Test Environment 4.0.0\")
 '    for each thing in MyFolder.files
 '	    response.write(thing&"<br />")
 '    next
'------------------------------------------------------------------------------
'Setup folder path
 lcl_file_directory = Application("userbids_upload_directory")
 'lcl_file_upload    = "egovlink300_admin\custom\pub"
 'lcl_file_folder1   = "\" & GetVirtualDirectyName()
 lcl_file_folder1   = GetVirtualDirectyName()
 lcl_file_folder2   = "\postings_bids"
 lcl_file_folder3   = "\userbids"

'if iorgid = 5 then
'   dtb_debug("1. file_directory: [" & lcl_file_directory & "] - file_upload: [" & lcl_file_upload & "] - file_folder1: [" & lcl_file_folder1 & "] - file_folder2: [" & lcl_file_folder2 & "] - file_folder3: [" & lcl_file_folder3 & "]")
'end if

'           lcl_blog_imgsrc = ""
'           lcl_blog_imgsrc = lcl_blog_imgsrc & Application("CommunityLink_DocUrl")
'           lcl_blog_imgsrc = lcl_blog_imgsrc & "/public_documents300/"
'           lcl_blog_imgsrc = lcl_blog_imgsrc & sorgVirtualSiteName
'           lcl_blog_imgsrc = lcl_blog_imgsrc & "/unpublished_documents"
'           lcl_blog_imgsrc = lcl_blog_imgsrc & oBlogInfo("imagefilename")



'Determine if the folder(s) have been created.  If not then create them for each level, if needed.
 checkPostingsUserBidsFolder lcl_file_directory & lcl_file_upload & lcl_file_folder1
 checkPostingsUserBidsFolder lcl_file_directory & lcl_file_upload & lcl_file_folder1 & lcl_file_folder2
 checkPostingsUserBidsFolder lcl_file_directory & lcl_file_upload & lcl_file_folder1 & lcl_file_folder2 & lcl_file_folder3

'Build the complete pathname
 uploadsDirVar        = lcl_file_directory & lcl_file_upload & lcl_file_folder1 & lcl_file_folder2 & lcl_file_folder3
 lcl_db_filedirectory = lcl_file_folder1 & lcl_file_folder2 & lcl_file_folder3

'Check to see if the user has attempted to upload the file.  If so then pull the filename entered
 lcl_upload_filename = request("upload_filename")
'dtb_debug("2. [" & uploadsDirVar & "] - [" & lcl_upload_filename & "]")
'Show screen status messages
 lcl_message  = ""
 lcl_filesize = 0
 lcl_onload   = ""

 if request("success") = "SU" then
    lcl_onload = "displayScreenMsg('*** Successfully Uploaded... ***');"
 elseif request("success") = "BIG" then
    lcl_filesize   = request("filesize")
    lcl_filesizeKB = (lcl_filesize/1024)    & " KB"
    lcl_filesizeMB = (lcl_filesize/1048576) & " MB"

    lcl_message = "<p><font style=\""color:#ff0000\"">"
    lcl_message = lcl_message & "<strong>*** ERROR: File Size too big. ***</strong><br />"
    lcl_message = lcl_message & "<font style=\""font-size:8pt\"">"
    'lcl_message = lcl_message & "<i>Must be under 10,240 KB or 10 MB</i><br />"
    lcl_message = lcl_message & "<i>Must be under 20,480 KB or 20 MB</i><br />"
    lcl_message = lcl_message & "Current file size: [" & lcl_filesizeKB & "] or [" & lcl_filesizeMB & "]"
    lcl_message = lcl_message & "</font><br />&nbsp;</p>"
 end if
%>
<html>
<head>
 	<title>E-Gov Services - <%=sOrgName%></title>

 	<link rel="stylesheet" type="text/css" href="css/styles.css" />
 	<link rel="stylesheet" type="text/css" href="global.css" />
 	<link rel="stylesheet" type="text/css" href="css/style_<%=iorgid%>.css" />

 	<script language="javascript" src="scripts/modules.js"></script>
 	<script language="javascript" src="scripts/easyform.js"></script>
  <script language="javascript" src="scripts/ajaxLib.js"></script>
  <script language="javascript" src="scripts/removespaces.js"></script>
  <script language="javascript" src="scripts/setfocus.js"></script>
  <script language="javascript" src="scripts/formvalidation_msgdisplay.js"></script>

<script language="javascript">
function onSubmitForm() {
  var lcl_success = "";
		var rege;
		var Ok;

  //Validate the Label
  if(document.getElementById("userLabel").value != "") {
     lcl_success = "Y";
     clearMsg('userLabel');
  }else{
     document.getElementById("userLabel").focus();
     inlineMsg(document.getElementById("userLabel").id,'<strong>Required Field Missing</strong> Label',10,'userLabel');
     lcl_success = "N";
  }

  //Validate the File
  if (document.getElementById("upload_filename").value == "") {
      document.getElementById("upload_filename").focus();
      inlineMsg(document.getElementById("upload_filename").id,'<strong>Required Field Missing</strong> File.<br />Please press the browse button and pick a file.',10,'upload_filename');
      lcl_success = "N";
  } else {
      lcl_filename = document.getElementById("upload_filename").value;
	 				rege         = /^[\w- :\\]+\.{1}[A-Za-z0-9]{2}[A-Za-z0-9]{0,2}$/;
		 			Ok           = rege.test(lcl_filename);

 					if(! Ok) {
         //lcl_invalid_msg  = '<strong>Invalid Value: </strong>';
         //lcl_invalid_msg += 'The filename has characters that are not allowed. ';
         //lcl_invalid_msg += 'Please rename the file on your PC prior to uploading it.<br />';
         //lcl_invalid_msg += '<strong>Example: </strong>My Doc_1-2006.txt';

         lcl_invalid_msg  = '<strong>Invalid Value: </strong>';
         lcl_invalid_msg += 'The path or filename has characters that are not allowed. ';
         lcl_invalid_msg += 'Please rename the file prior to uploading it. ';
         lcl_invalid_msg += '<strong>Allowed Characters: </strong>A through Z, a through z, 0 through 9, ';
         lcl_invalid_msg += 'underscore, dash, spaces, and one period before the file extension.<br />';
         lcl_invalid_msg += '<strong>Example: </strong>C:\\My Documents\\Bid Postings\\My Doc_1-2206.txt';

         document.getElementById("upload_filename").focus();
         inlineMsg(document.getElementById("upload_filename").id,lcl_invalid_msg,10,'upload_filename');
         lcl_success = "N";
					 }	else {
         lcl_length     = lcl_filename.length;
         lcl_period_loc = lcl_filename.indexOf(".");
         lcl_ext        = lcl_filename.substr(lcl_period_loc+1)
       <%
        'File extensions allowed
         lcl_display_ext = "BMP, DOC, GIF, JPG, PDF, TXT, XLS, or ZIP "

        'Set up the javascript checks for each file extension
         lcl_ext_allowed = "(lcl_ext.toUpperCase()==""BMP"")"
         lcl_ext_allowed = lcl_ext_allowed & "||(lcl_ext.toUpperCase()==""DOC"")"
         lcl_ext_allowed = lcl_ext_allowed & "||(lcl_ext.toUpperCase()==""GIF"")"
         lcl_ext_allowed = lcl_ext_allowed & "||(lcl_ext.toUpperCase()==""JPG"")"
         lcl_ext_allowed = lcl_ext_allowed & "||(lcl_ext.toUpperCase()==""PDF"")"
         lcl_ext_allowed = lcl_ext_allowed & "||(lcl_ext.toUpperCase()==""TXT"")"
         lcl_ext_allowed = lcl_ext_allowed & "||(lcl_ext.toUpperCase()==""XLS"")"
         lcl_ext_allowed = lcl_ext_allowed & "||(lcl_ext.toUpperCase()==""ZIP"")"
       %>
         if(<%=lcl_ext_allowed%>) {
            lcl_filename = document.getElementById("upload_filename").value;
            document.getElementById("postings_submit_userbids").action="postings_submit_userbids.asp?posting_id=<%=lcl_posting_id%>&dlistid=<%=lcl_dlistid%>&listtype=<%=lcl_list_type%>&upload_filename="+lcl_filename;
            clearMsg('upload_filename');
            if(lcl_success!="N") {
               lcl_success = "Y";
            }
         }else{
            lcl_msg  = "<strong>Invalid Value: </strong>File Type (" + lcl_ext + "). ";
            lcl_msg += "Valid file types include: <strong><%=lcl_display_ext%></strong>";

            document.getElementById("upload_filename").focus();
            inlineMsg(document.getElementById("upload_filename").id,lcl_msg,10,'upload_filename');
            lcl_success = "N";

            //document.getElementById("postings_submit_userbids").action="postings_submit_userbids.asp?posting_id=<%=lcl_posting_id%>&dlistid=<%=lcl_dlistid%>&listtype=<%=lcl_list_type%>&upload_filename="+lcl_filename;
            //displayScreenMsg(lcl_msg);
            //lcl_success = "N";
         }
      }
  }

  if(lcl_success=="Y") {
     showUploadingMsg();
     document.getElementById("postings_submit_userbids").submit();
  }
}

function showUploadingMsg() {
  lcl_msg  = "*** Uploading... ***<br />";
  lcl_msg += "<font style=\"font-size:8pt; color:#ff0000\">";
  lcl_msg += "Depending on the size of the file<br />";
  lcl_msg += "this may take several minutes to complete.<br />";
  //lcl_msg += "<i>(10 MB or 10,240 KB file size MAX)</i>";
  lcl_msg += "<i>(20 MB or 20,480 KB file size MAX)</i>";
  lcl_msg += "</font>";

  document.getElementById("screenMsg").innerHTML = lcl_msg;
}

function displayScreenMsg(iMsg) {
  if(iMsg!="") {
     document.getElementById("screenMsg").innerHTML = iMsg;
     window.setTimeout("clearScreenMsg()", (10 * 1000));
  }
}

function clearScreenMsg() {
  document.getElementById("screenMsg").innerHTML = "";
}
</script>
</head>
<%
 response.write "<body onload=""document.getElementById('upload_filename').focus();" & lcl_onload & """>" & vbcrlf

'Determine if the user is attempting an upload or not
 Dim diagnostics
 if Request.ServerVariables("REQUEST_METHOD") <> "POST" then
    diagnostics = TestEnvironment()
    if diagnostics<>"" then
       response.write "<div style=""margin-left:20; margin-top:30; margin-right:30; margin-bottom:30;"">" & vbcrlf
       response.write diagnostics
       response.write "<p>After you correct this problem, reload the page." & vbcrlf
       response.write "</div>" & vbcrlf
    else
       OutputForm()
    end if
 else
    checkPostingsUserBidsFolder uploadsDirVar
    uploadUserBidFile()
 end if
%>
  </div>
</div>
</body>
</html>
<%
set oOrg = nothing

'------------------------------------------------------------------------------
function OutputForm()
%>
<table border="0" cellspacing="0" cellpadding="2" width="340">
  <form name="postings_submit_userbids" id="postings_submit_userbids" method="post" enctype="multipart/form-data" action="">
  <caption align="center" valign="top">
    <p><span id="screenMsg" style="color:#ff0000;font-weight:bold;font-size:10pt;"></span><br />&nbsp;</p>
  </caption>
  <tr>
      <td>
          <fieldset>
            <legend><font class="pagetitle"><%=lcl_feature_name%>: Upload Bid</font>&nbsp;</legend>
          <table border="0" cellspacing="0" cellpadding="2" width="100%">
            <tr>
                <td><strong>File:</strong></td>
                <td><input type="file" name="upload_filename" id="upload_filename" value="" size="50" maxlength="1000" onchange="clearMsg('upload_filename');" /></td>
            </tr>
            <tr>
                <td><strong>Label:</strong></td>
                <td><input type="text" name="userLabel" id="userLabel" size="50" maxlength="500" onchange="clearMsg('userLabel');" /></td>
            </tr>
            <tr>
                <td colspan="2" align="center">
                    <input type="button" name="sAction" value="Upload Bid" class="button" onclick="onSubmitForm();" />
                    <input type="button" name="closewindow" value="Cancel" class="button" onclick="parent.close()" />
                </td>
            </tr>
          </table>
          </fieldset>
      </td>
  </tr>
  </form>
</table>
<%
 'Display the screen message is one exists.
  if lcl_message <> "" then
     response.write "<script language=""javascript"">" & vbcrlf
     response.write "  displayScreenMsg('" & lcl_message & "');" & vbcrlf
     response.write "</script>" & vbcrlf
  end if

end function

'------------------------------------------------------------------------------
function TestEnvironment()
    Dim fso, mf, fileName, testFile, streamTest
    TestEnvironment = ""
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    if not fso.FolderExists(uploadsDirVar) then
       set mf = fso.CreateFolder(uploadsDirVar)
    end if
    fileName = uploadsDirVar & "\test.txt"
    on error resume next
    Set testFile = fso.CreateTextFile(fileName, true)
    If Err.Number<>0 then
        TestEnvironment = "<strong>Folder " & uploadsDirVar & " does not have write permissions.</strong><br />The value of your uploadsDirVar is incorrect. Open uploadTester.asp in an editor and change the value of uploadsDirVar to the pathname of a directory with write permissions."
        exit function
    end if
    Err.Clear
    testFile.Close
    fso.DeleteFile(fileName)
    If Err.Number<>0 then
        TestEnvironment = "<strong>Folder " & uploadsDirVar & " does not have delete permissions</strong>, although it does have write permissions.<br />Change the permissions for IUSR_<I>computername</I> on this folder."
        exit function
    end if
    Err.Clear
    Set streamTest = Server.CreateObject("ADODB.Stream")
    If Err.Number<>0 then
        TestEnvironment = "<strong>The ADODB object <I>Stream</I> is not available in your server.</strong><br />Check the Requirements page for information about upgrading your ADODB libraries."
        exit function
    end if
    Set streamTest = Nothing
end function

'------------------------------------------------------------------------------
sub checkPostingsUserBidsFolder(sFolderPath)

	set oFSO = server.createobject("Scripting.FileSystemObject")
	
	if oFSO.FolderExists(sFolderPath) <> True then
 		'Create postings/userbids folder
  		set oFolder = oFSO.CreateFolder(sFolderPath)
		  set oFolder = nothing
 end if

	set oFSO = nothing

end sub

'------------------------------------------------------------------------------
sub uploadUserBidFile()

 dim oUpload

'Create the upload object
 set oUpload = Server.CreateObject("Dundas.Upload.2")
'The "MaxFileSize" is set so high because if the file size is greater than this then the script bombs.
 oUpload.MaxFileSize = (41943040 * 5) ' MAX SIZE OF UPLOAD SPECIFIED IN BYTES, (4096000 * 5) =  APPX. 200MB
 oUpload.SaveToMemory

'Get the file to be uploaded
 lcl_uploadfile = oUpload.Files(0).OriginalPath

'------------------------------------------------------------------------------
'Determine if the file is too big or not (20 MB)
' 1 Byte     = 8 Bit
' 1 Kilobyte = 1024 Bytes
' 1 Megabyte = 1048576 Bytes
' 1 Gigabyte = 1073741824 Bytes
'------------------------------------------------------------------------------
 'if CLng(oUpload.Files(0).Size) <= 10485760 then 'APPX. 10 MB
 if CLng(oUpload.Files(0).Size) <= 20981520 then 'APPX. 20 MB

   'Set the variables
    sFileName = LCASE(RIGHT(lcl_uploadfile,LEN(lcl_uploadfile) - instrrev(lcl_uploadfile,"\")))
    iRequestId   = oUpload.Form("iRequestId")
    iAdminUserId = session("UserID")
    sPostingID   = oUpload.form("posting_id")
    sListType    = oUpload.form("listtype")
    sdlistid     = oUpload.form("dlistid")
    sUserLabel   = oUpload.form("userLabel")

    'sServerPath = server.mappath("../") & "\egovlink300_admin\custom\pub\" & sorgVirtualSiteName & "\postings\user_bids"

    'sServerPath = server.mappath("../")
    'sFolder1 = "\egovlink300_admin\custom\pub\" & sorgVirtualSiteName
    'sFolder2 = "\postings_bids"
    'sFolder3 = "\userbids"

    sServerPath = Application("userbids_upload_directory")
    sFolder1    = "\" & GetVirtualDirectyName()
    sFolder2    = "\postings_bids"
    sFolder3    = "\userbids"

   'Check to see if the folder exists - Create it if it does not
    checkPostingsUserBidsFolder sServerPath & sFolder1
    checkPostingsUserBidsFolder sServerPath & sFolder1 & sFolder2
    checkPostingsUserBidsFolder sServerPath & sFolder1 & sFolder2 & sFolder3

    lcl_filepath = sServerPath & sFolder1 & sFolder2 & sFolder3

   'Store the file in the server filesystem
    if oUpload.FileExists( lcl_filepath & "\" & sFileName ) then
 	    'Delete the file if it already exists on server filesystem
      	oUpload.FileDelete( lcl_filepath & "\" & sFileName )
    end if

   'Save file on server filesystem
    oUpload.Files(0).SaveAs(  lcl_filepath & "\" & sFileName )

   'Create the userbids record and send out emails
    SaveFiles sUserLabel

    lcl_success = "SU"

 else
    lcl_success = "BIG&filesize=" & oUpload.Files(0).Size
 end if

 set oUpload = nothing

 response.redirect "postings_submit_userbids.asp?posting_id=" & lcl_posting_id & "&listtype=" & lcl_list_type & "&dlistid=" & lcl_dlistid & "&success=" & lcl_success

end sub

'------------------------------------------------------------------------------
sub SaveFiles(iUserLabel)

 'Create the egov_jobs_bids_userbids record
  'storePostingsUserBidsInfo lcl_posting_id, lcl_list_type, uploadsDirVar, lcl_upload_filename, iUserLabel, lcl_uploadid
  storePostingsUserBidsInfo lcl_posting_id, lcl_list_type, lcl_db_filedirectory, lcl_upload_filename, iUserLabel, lcl_uploadid

 'Get the user email and send upload email
  lcl_useremail = getUserEmail(request.cookies("userid"))
  sendUploadEmail lcl_posting_id, lcl_useremail, "", uploadsDirVal, lcl_upload_filename, lcl_uploadid, iUserLabel

 'Get the admin email and send upload email
  lcl_orgemail = getPostingsAdminEmail()
  sendUploadEmail lcl_posting_id, "", lcl_orgemail, uploadsDirVal, lcl_upload_filename, lcl_uploadid, iUserLabel

end sub

'------------------------------------------------------------------------------
sub storePostingsUserBidsInfo(ByVal ipostingid, ByVal ipostingtype, ByVal ifilelocation, ByVal iuploadfilename, _
                              ByVal iUserLabel, ByRef lcl_uploadid)
  if dbready_number(ipostingid) then

     lcl_postingtype    = dbready_string(ipostingtype,50)
     lcl_filelocation   = dbready_string(ifilelocation,1000)
     lcl_uploadfilename = dbready_string(iuploadfilename,500)
     lcl_userlabel      = dbready_string(iUserLabel,500)

    'Build message for upload notification.
     lcl_upload_message = ""
     lcl_upload_message = lcl_upload_message & "orgid: ["              & iorgid             & "]<br />" & vbcrlf
     lcl_upload_message = lcl_upload_message & "Date: ["               & now()              & "]<br />" & vbcrlf
     lcl_upload_message = lcl_upload_message & "lcl_filelocation: ["   & lcl_filelocation   & "]<br />" & vbcrlf
     lcl_upload_message = lcl_upload_message & "lcl_uploadfilename: [" & lcl_uploadfilename & "]<br />" & vbcrlf

    'Format filelocation
     if lcl_filelocation <> "" then
        lcl_filelocation = replace(lcl_filelocation,lcl_file_directory & "egovlink300_admin","")
     end if

    'Format filename
     if lcl_uploadfilename <> "" then
        lcl_uploadfilename    = replace(lcl_uploadfilename,lcl_file_directory & "\egovlink300_admin","")
        lcl_filename_startloc = InStrRev(lcl_uploadfilename,"\")

        if lcl_filename_startloc > 0 then
           lcl_uploadfilename = mid(lcl_uploadfilename,lcl_filename_startloc)
        end if

        if left(lcl_uploadfilename,1) <> "\" then
           lcl_uploadfilename = "\" & lcl_uploadfilename
        end if

       'Setup upload message to DTB
        lcl_upload_message = lcl_upload_message & "left: [" & left(lcl_uploadfilename,1) & "]<br />"
        lcl_upload_message = lcl_upload_message & "fileuploaded: [" & lcl_uploadfilename & "]"

        if left(lcl_uploadfilename,1) <> "\" then
           sendEmail "dboyer@eclink.com", "dboyer@eclink.com","","User Bids: Uploaded File - Missing ""\""",lcl_upload_message,"","Y"
           'dtb_debug(lcl_upload_message)
        end if

     end if

'dtb_debug("3. [" & lcl_uploadfilename & "]")

    'Insert the userbid
     sSQLi = "INSERT INTO egov_jobs_bids_userbids ("
     sSQLi = sSQLi & "posting_id, "
     sSQLi = sSQLi & "posting_type, "
     sSQLi = sSQLi & "userid, "
     sSQLi = sSQLi & "orgid, "
     sSQLi = sSQLi & "submitdate, "
     sSQLi = sSQLi & "filelocation, "
     sSQLi = sSQLi & "filename, "
     sSQLi = sSQLi & "userLabel "
     sSQLi = sSQLi & ") VALUES ("
     sSQLi = sSQLi &       ipostingid                & ", "
     sSQLi = sSQLi & "'" & lcl_postingtype           & "', "
     sSQLi = sSQLi &       request.cookies("userid") & ", "
     sSQLi = sSQLi &       iorgid                    & ", "
     sSQLi = sSQLi & "'" & now()                     & "', "
     sSQLi = sSQLi & "'" & lcl_filelocation          & "', "
     sSQLi = sSQLi & "'" & lcl_uploadfilename        & "', "
     sSQLi = sSQLi & "'" & lcl_userlabel             & "' "
     sSQLi = sSQLi & ")"

     set rsi = Server.CreateObject("ADODB.Recordset")
     rsi.Open sSQLi, Application("DSN"), 0, 1

    'Retrieve the posting_id that was just inserted
     sSQLid = "SELECT IDENT_CURRENT('egov_jobs_bids_userbids') as NewID"
     rsi.Open sSQLid, Application("DSN"), 3, 1
     lcl_identity = rsi.Fields("NewID").value

    'Create the uploadid
     lcl_uploadid = createUploadID(request.cookies("userid"), lcl_identity)

    	set rsi = nothing

  end if

end sub

'------------------------------------------------------------------------------
function createUploadID(p_userid, p_identity)
  lcl_return   = ""
  lcl_uploadid = ""

  if p_identity <> "" AND p_userid <> "" then
     lcl_uploadid = "BID" & year(now()) & month(now()) & day(now()) & p_userid & p_identity

     sSQLu = "UPDATE egov_jobs_bids_userbids SET uploadid = '" & lcl_uploadid & "' WHERE userbidid = " & p_identity
     set rsu = Server.CreateObject("ADODB.Recordset")
     rsu.Open sSQLu, Application("DSN"), 0, 1

     set rsu = nothing

     lcl_return = lcl_uploadid

  end if

  createUploadID = lcl_return

end function

'------------------------------------------------------------------------------
sub dtb_debug(p_value)
  sSQLi = "INSERT INTO my_table_dtb(notes) VALUES('" & replace(p_value,"'","''") & "')"
  set dtb = Server.CreateObject("ADODB.Recordset")
  dtb.Open sSQLi, Application("DSN"), 0, 1

end sub
%>