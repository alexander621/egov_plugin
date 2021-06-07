<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="image_upload_new.asp" -->
<%
'Check to see if the feature is offline
 if isFeatureOffline("memberships") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

'option explicit 
 Response.Expires     = -1
 Server.ScriptTimeout = 600

 sLevel = "../"  'Override of value from common.asp

'Determine if this is a demo or not.  demo = Y means that these screens can function without the web camera attached
 lcl_demo = request("demo")

'Set up demo variables
 lcl_demo_page_title = ""
 lcl_demo_url        = ""

 if lcl_demo = "Y" then
    lcl_demo_page_title = " (DEMO)"
    lcl_demo_url        = "&demo=" & lcl_demo
 end if

'Retreive the MemeberID
 dim lcl_member_id, lcl_reload_pic, lcl_file_directory
 lcl_member_id  = request("memberid")
 lcl_reload_pic = request("reload_pic")
 lcl_poolpassid = request("poolpassid")

'------------------------------------------------------------------------------
'Display all of the files in the directory *** FOR DEBUGGING ONLY ***
 'set MyFile   = server.createobject("scripting.filesystemobject")
 'set MyFolder = MyFile.GetFolder("d:\wwwroot\www.cityegov.com\egovlink300_admin\")
 'set MyFolder = MyFile.GetFolder("d:\wwwroot\www.cityegov.com\egovlink release QA Test Environment 4.0.0\")
 '    for each thing in MyFolder.files
 '	    response.write(thing&"<br />")
 '    next

'------------------------------------------------------------------------------
'Depending on the environment the folder will need to be set differently.
'You will need to uncomment out the one you need.
'sSystem = "DVLP"

'Select Case sSystem 
'	Case "DVLP"
'		lcl_file_directory = "c:\www_server_root\egovlink\egovlink release 4.0.0\"
'	Case "TEST"
'      lcl_file_directory = "c:\wwwroot\www.cityegov.com\egovlink release QA Test Environment 4.0.0\"
'	Case "PROD"
'      lcl_file_directory = "c:\wwwroot\www.cityegov.com\"
'	Case Else
'		    lcl_file_directory = "c:\www_server_root\egovlink\egovlink release 4.0.0\"
'End Select
lcl_file_directory = Application("membershipcard_filedirectory")

'If this is a demo then no actual picture will need to be taken.
'Pull image from the DEMO folder.
 if lcl_demo = "Y" then
    'lcl_file_upload = "egovlink300_admin\images\MembershipCard_Photos\demo"
    lcl_file_upload = "\demo"
 else
    'lcl_file_upload = "egovlink300_admin\images\MembershipCard_Photos\temp"
    lcl_file_upload = "\temp"
 end if

lcl_file_directory = lcl_file_directory & lcl_file_upload

'------------------------------------------------------------------------------
' Change the value of the variable below to the pathname
' of a directory with write permissions, for example "C:\Inetpub\wwwroot"
  Dim uploadsDirVar

  'uploadsDirVar = "d:\inetpub\webmailasp\database\tempUploads" 
  uploadsDirVar = lcl_file_directory
'------------------------------------------------------------------------------
%>
<html>
<head>
  <title>E-Gov Administration Console {Membership Photo Taking System}</title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />

<style type="text/css">
  #returnButton {
     margin-bottom: 10px;
  }

  #button1 {
     left:   17px;
     width:  150px;
     top:    363px;
     height: 29px
  }

  #diagnostics {
     margin-left:   20;
     margin-top:    30;
     margin-right:  30;
     margin-bottom: 30;
  }

  #file_name,
  #upload_image {
     width: 380px;
  }

  #displayFileName {
    text-align:  center;
    font-size:   20px;
    color:       #ff0000;
    font-weight: bold;
  }

  #uploadImageButton {
     margin-top: 4px;
     left:       17px;
     width:      150px;
     top:        363px;
     height:     29px
  }
</style>

<script language="Javascript">
// Register the Twain interface
function register() {
  var User;
  var Domain;
  var RegCode;

  //User    = "Jerry Felix";
  //Domain  = "egovlink.com";
  //RegCode = "74F85955BBA574EBB983FE2F8EF479EB928CABF18C14ACB8BEC7715B484E95C86BB464D1151069EFBF19BD3094963A4642E78EA1EAD55CC8AFF321B09B62016CFF1AC3EA30FC03413FA72DCBD774E009ABA053A5151405F0355522C0277E053C23868C0CED1DE34766CD5C4EFCBA3EDDB1C6D3EF3685853973CE69067FCCE9EA407A1537D47A04CD6D62398FC2A9AE823E01ABECF8FEE61B443D50FCF443EFB9";

  User    = "EC Link";
  Domain  = "www.egovlink.com";
  RegCode = 'A0B5582BA4CC988742C3920E996E5B5EC2FA83BBEDC487A036A6F8C505E4934700E4FE3862EE8E3EECE001B52B7D508BF97194CE0D83D8C635A9D8588126C02DFD69A8DFCEF35710F138571492E78858FF9D5391AC07999BD8DEE5B58DD494DC71016088917E0C25F9C6CF9E53EB823F977ABC216FF3C08A5E2FC37DB2478352638FB1F18725CAD94A843E96410196A7A61887C49AD6119C474BC0FCBF34E52D';

  VSTwain1.Register( User, Domain, RegCode );
}

// Setup the Camera

var previewImagePath

function init() {
  // hide these fields when the page is first opened.
  //document.getElementById("file_name").style.display     = "none";
//  document.getElementById("upload_image").style.display  = "none";

  // set the color to RED for the first instruction with the page is first opened.
  document.getElementById("instruction_1").style.color   = "#800000";
  document.getElementById("instruction_2").style.color   = "#000000";
  document.getElementById("instruction_3").style.color   = "#000000";

//  previewImagePath = VSTwain1.GetPathToTempDir() + "test.jpg";
}

// Start the camera program and take the picture
function WithDialog() {
  //  init();

  //Set up the registration
  VSTwain1.StartDevice();
  register();
  VSTwain1.MaxImages           = 1;
  VSTwain1.AutoCleanBuffer     = 1;
  VSTwain1.object.SourceIndex  = 0;
  //alert(VSTwain1.SourcesCount);
  //alert('is twain available: [' + VSTwain1.IsTwainAvailable + ']');
  //VSTwain1.TwainDllPath = 'C:\Windows\System32\';

  //alert('twain dll path: [' + VSTwain1.TwainDllPath + ']');

  //if(VSTwain1.SelectSource() == 1) {
     VSTwain1.ShowUI              = 1;
     VSTwain1.DisableAfterAcquire = 1;
     VSTwain1.OpenDataSource();
     VSTwain1.PixelType           = 2;  // RGB
     VSTwain1.UnitOfMeasure       = 0;  // inches
     //VSTwain1.Resolution          = 600;
  //}

  VSTwain1.Acquire()
  //VSTwain1.ShowUI = 1;
  //VSTwain1.Acquire();
  //VSTwain1.DeleteImage(VSTwain1.numImages)
}

// Stop the Twain interface
function onPageUnload() {
  VSTwain1.StopDevice()
}

function GoBack() {
  var lcl_return_page = '<%=session("redirectpage")%>';

  if (lcl_return_page != "") {
      location.href = lcl_return_page;
  } else {
      history.go(-1);
  }
}

function onSubmitForm() {
    var formDOMObj = document.frmSend;
    if (formDOMObj.attach1.value == "")
        alert("Please press the browse button and pick a file.")
    else
        return true;
    return false;
}

function check_instruction3() {
  if(document.frmSend.attach1.value != "") {
     document.getElementById("instruction_2").style.color  = "#000000";
     document.getElementById("instruction_3").style.color  = "#800000";
//     document.getElementById("upload_image").style.display = "block";
  }else{
     document.getElementById("instruction_2").style.color  = "#800000";
     document.getElementById("instruction_3").style.color  = "#000000";
//     document.getElementById("upload_image").style.display = "none";
  }

}

function openWin(page) {
  OpenWin = window.open(page, "new", "toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes,width=505,height=370,screenX=0,screenY=0");
  if (document.images) {OpenWin.focus();}
}
</script>

<script language="javascript" event="PostScan(flag)" for="VSTwain1">
<!--
     if(flag != 0) {
        if (VSTwain1.errorCode != 0) alert(VSTwain1.errorString)
     } else {
        var imgPath = "c:/MembershipCard_Photos/<%=lcl_member_id%>.jpg";
        //var imgPath = "c:/MembershipCard_Photos/<%'lcl_member_id%>.bmp";

        if(VSTwain1.SaveImage(0,imgPath) == 0) {
           alert(VSTwain1.errorString + ' (error code: ' + VSTwain1.errorCode + ')');
           alert("Please try to scan black-white or gray images!");
        }else{
           var img = new Image();
           img.src = imgPath;

      		   document.getElementById("file_name").style.display     = "block";
           document.getElementById("instruction_2").style.display = "block";

           document.getElementById("instruction_1").style.color   = "#000000";
           document.getElementById("instruction_2").style.color   = "#800000";

           document.getElementById("button1").value               = "Retake Picture";
           document.getElementById("path_name").innerHTML         = "<font style=\"font-size: 14px; font-style: italic; \">The file can be found here: </font><font style=\"font-size: 20px; color: FF0000; font-weight: bold; \">" + imgPath + "</font>";
        }
    }
//     VSTwain1.CloseDataSource()
// -->
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onUnload="onPageUnload()" onLoad="init();">
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
 'BEGIN: Camera Object --------------------------------------------------------
  'response.write "<object id=""VSTwain1"" width=""1"" height=""1"" classid=""CLSID:1169E0CD-9E76-11D7-B1D8-FB63945DE96D"" codebase=""camera_dvlp_driver/VSTwain_new.dll""></object>" & vbcrlf
'  response.write "<object id=""VSTwain1"" width=""1"" height=""1"" classid=""CLSID:1169E0CD-9E76-11D7-B1D8-FB63945DE96D"" codebase=""camera_dvlp_driver_new/VSTwain.dll""></object>" & vbcrlf
  response.write "<object id=""VSTwain1"" width=""1"" height=""1"" classid=""CLSID:1169E0CD-9E76-11D7-B1D8-FB63945DE96D"" codebase=""""></object>" & vbcrlf
 'END: Camera Object ----------------------------------------------------------

  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf
  response.write "<h3>Taking a Picture" & lcl_demo_page_title & "</h3>" & vbcrlf
  response.write "<input type=""button"" name=""returnButton"" id=""returnButton"" value=""" & session("RedirectLang") & """ class=""button"" onclick=""GoBack();"" />" & vbcrlf
  response.write "<p>" & vbcrlf
  response.write "  <strong>Instructions: </strong><br />" & vbcrlf

 'BEGIN: Instruction 1 --------------------------------------------------------
 'If this is a demo then simulate the picture being taken
  if lcl_demo = "Y" then
     lcl_onclick = "openWin('demo_image_takepic.asp?memberid=" & lcl_member_id & "&poolpassid=" & lcl_poolpassid & "&demo=Y')"
  else
     lcl_onclick = "WithDialog();"
  end if

  response.write "  <span id=""instruction_1"">&nbsp;&nbsp;<strong>1. </strong>Click the ""Take Picture"" button when you are ready to take the member's picture.<p></span>" & vbcrlf
  response.write "  <table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
  response.write "    <tr><td><input type=""button"" name=""button1"" id=""button1"" onclick=""" & lcl_onclick & """ value=""Take Picture"" /></td></tr>" & vbcrlf
  response.write "  </table>" & vbcrlf
  response.write "<p>" & vbcrlf

 'Set up the file input form field
  Dim diagnostics
  if Request.ServerVariables("REQUEST_METHOD") <> "POST" then
     diagnostics = TestEnvironment()

     if diagnostics <> "" then
        response.write "<div id=""diagnostics"">" & vbcrlf
        response.write    diagnostics
        response.write "  <p>After you correct this problem, reload the page." & vbcrlf
        response.write "</div>" & vbcrlf
     else
        response.write "<div>" & vbcrlf
                          OutputForm()
        response.write "</div>" & vbcrlf
     end if
  else
     response.write "<div>" & vbcrlf
                       OutputForm()
     response.write    SaveFiles()
     response.write "  <br /><br />" & vbcrlf
     response.write "</div>" & vbcrlf
  end if

  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"-->  
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
function OutputForm()

  response.write "<form name=""frmSend"" id=""frmSend"" method=""POST"" enctype=""multipart/form-data"" action=""image_takepic_new.asp?memberid=" & lcl_member_id & "&poolpassid=" & lcl_poolpassid & "&demo=" & lcl_demo & """ onsubmit=""return onSubmitForm();"">" & vbcrlf

 'BEGIN: Instruction 2 --------------------------------------------------------
  response.write "    <span id=""instruction_2"">&nbsp;&nbsp;<strong>2. </strong>Once the picture has been taken you must now upload the file.  To do this you must enter the filename of " & vbcrlf
  response.write "        the image into the ""File Name"" field.  You can do this by either typing in the value or by clicking on the ""Browse"" button and selecting the file " & vbcrlf
  response.write "        itself.  The file you need is identified under the ""Browse"" button.</span>" & vbcrlf
  response.write "    <p>" & vbcrlf
  response.write "    <table id=""file_name"" border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
  response.write "      <tr>" & vbcrlf
  response.write "          <td>" & vbcrlf

  if request("step2") = "Y" then
     response.write "<p>" & vbcrlf
     'response.write "<strong>File name:</strong>&nbsp;<input name=""attach1"" type=""text"" value=""c:/MembershipCard_Photos/" & lcl_member_id & ".jpg"" size=""35"">&nbsp;" & vbcrlf
     response.write "  <strong>File name:</strong>&nbsp;<input name=""attach1"" type=""text"" value=""c:/MembershipCard_Photos/" & lcl_member_id & ".jpg"" size=""35"" />&nbsp;" & vbcrlf
     response.write "  <input type=""button"" value=""Browse"" /><br />" & vbcrlf
     response.write "  <div id=""displayFileName"">c:/MembershipCard_Photos/" & lcl_member_id & ".jpg</div>" & vbcrlf
     'response.write "<div align=""center""><font style=""font-size:20px; color:#FF0000; font-weight:bold;"">c:/MembershipCard_Photos/" & lcl_member_id & ".bmp</font></div>" & vbcrlf
     response.write "</p>" & vbcrlf
  else
     response.write "<p>" & vbcrlf
     response.write "  <strong>File name:</strong>&nbsp;<input name=""attach1"" type=""file"" size=""35"" onChange=""check_instruction3()"" /><br />" & vbcrlf
     response.write "  <div align=""center"" id=""path_name"">&nbsp;</div>" & vbcrlf
     response.write "</p>" & vbcrlf
  end if

  response.write "          </td>" & vbcrlf
  response.write "      </tr>" & vbcrlf
  response.write "    </table>" & vbcrlf
 'END: Instruction 2 ----------------------------------------------------------

 'BEGIN: Instruction 3 --------------------------------------------------------
  response.write "    <span id=""instruction_3"">&nbsp;&nbsp;<strong>3. </strong>Once you have selected the image to upload, you can then click on the ""Upload Image"" button." & vbcrlf
  response.write "      This will set up the member's ID to be printed.<p></span>" & vbcrlf
  response.write "    <table id=""upload_image"" border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf

  if lcl_demo = "Y" then
     response.write "      <tr><td><input type=""button"" name=""uploadImageButton"" id=""uploadImageButton"" value=""Upload Image"" onclick=""javascript:location.href='image_display_new.asp?memberid=" & lcl_member_id & "&poolpassid=" & lcl_poolpassid & "&demo=Y'"" /></td></tr>" & vbcrlf
  else
     response.write "      <tr><td><input type=""submit"" name=""uploadImageButton"" id=""uploadImageButton"" value=""Upload Image"" /></td></tr>" & vbcrlf
  end if

  response.write "</table>" & vbcrlf
 'END: Instruction 3 ----------------------------------------------------------

  response.write "</form>" & vbcrlf

end function

'------------------------------------------------------------------------------
function TestEnvironment()
    Dim fso, mf, fileName, testFile, streamTest
    TestEnvironment = ""
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    if not fso.FolderExists(uploadsDirVar) then
       set mf = fso.CreateFolder(lcl_file_directory)
'        TestEnvironment = "<strong>Folder " & uploadsDirVar & " does not exist.</strong><br />The value of your uploadsDirVar is incorrect. Open uploadTester.asp in an editor and change the value of uploadsDirVar to the pathname of a directory with write permissions."
'        exit function
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
function SaveFiles()
    Dim Upload, fileName, fileSize, ks, i, fileKey

'	Set Upload = New FreeASPUpload
	Set Upload = New ImageUpload
	'response.write uploadsDirVar
	'response.end
 Upload.Save(uploadsDirVar)

	' If something fails inside the script, but the exception is handled
	If Err.Number<>0 then Exit function

    SaveFiles = ""
    ks = Upload.UploadedFiles.keys
    if (UBound(ks) <> -1) then
  	    Session("RELOAD_PIC") = "Y"
       response.redirect "image_display_new.asp?memberid=" & lcl_member_id & "&poolpassid=" & lcl_poolpassid

'        SaveFiles = "<strong>Files uploaded:</strong> "
'        for each fileKey in Upload.UploadedFiles.keys
'            SaveFiles = SaveFiles & Upload.UploadedFiles(fileKey).FileName & " (" & Upload.UploadedFiles(fileKey).Length & "B) "
'        next
    else
        SaveFiles = "The file name specified in the upload form does not correspond to a valid file in the system."
    end if
'	SaveFiles = SaveFiles & "<br />Enter a number = " & Upload.Form("enter_a_number") & "<br />"
'	SaveFiles = SaveFiles & "Checkbox values = " & Upload.Form("checkbox_values") & "<br />"
end function
%>
