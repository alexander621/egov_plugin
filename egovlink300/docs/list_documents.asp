<!-- #include file="../includes/common.asp" //-->


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
	<TITLE> New Document </TITLE>
	<META NAME="Generator" CONTENT="EditPlus">
	<META NAME="Author" CONTENT="">
	<META NAME="Keywords" CONTENT="">
	<META NAME="Description" CONTENT="">
	<link rel="stylesheet" href="../css/style_<%=iorgid%>.css" type="text/css">
	<script language="JavaScript1.2" src="../scripts/doctreenav.js"></script>
	<script language="JavaScript1.2" src="../scripts/docfolderopen.js" ></script>
</HEAD>

<BODY>


<!--#Include file="../include_top.asp"-->


<!--BEGIN: DISPLAY DOCUMENT/FOLDER TREE-->
<%
' GENERATE DOCUMENT/FOLDER TREE
sLocationName =  GetVirtualDirectyName()

Set FSO = CreateObject("Scripting.FileSystemObject")

response.write "<ul>" & vbcrlf

Call ENUMERATEDOCUMENTS(FSO.GetFolder(Server.Mappath("/public_documents300/" & sLocationName & "/published_documents/")),Server.Mappath("/public_documents300/" & sLocationName & "/published_documents/"),"public_documents300/" & sLocationName & "/published_documents"  )

response.write "</uL>"

' DESTROY OBJECTS
Set FSO = Nothing
%>
<!--END: DISPLAY DOCUMENT/FOLDER TREE-->


<!--BEGIN: CODE TO OPEN SELECTED FOLDER FROM EXTERNAL LINK IF FOLDER ID SUPPLIED-->
<% If request("egovlinkfid") <> "" Then %>
	<script language="JavaScript1.2">
		OpenFolder(<%=clng(request("egovlinkfid"))%>);
	</script>
<% End If %>
<!--END: CODE TO OPEN SELECTED FOLDER FROM EXTERNAL LINK-->


</BODY>
</HTML>

<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------


'-------------------------------------------------------------------------------------------------------
' SUB ENUMERATEDOCUMENTS(FOLDER)
'-------------------------------------------------------------------------------------------------------
Sub ENUMERATEDOCUMENTS(Folder,sPath,sVpath)
     
	' BUILD HYPERLINK BASE PATH
	sVirtualpath = replace(Folder.Path,sPath,"")
	sVirtualpath = "http://www.egovlink.com/" & sVpath & replace(sVirtualpath,"\","/") 


		' LIST CONTENTS OF FOLDER (SUBFOLDERS AND FILES)
		For Each SubFolder in Folder.SubFolders
		
			' WRITE FOLDER INFORMATION
			response.write vbcrlf
			response.write "<li id=""foldheader"">" & SubFolder.Name & "</LI>" & vbcrlf
			response.write "<ul id=""foldinglist"" style=""display:none"" >" & vbcrlf
		
			' RECURSIVE CALL TO GET ANY SUBFOLDERS OF THE CURRENT FOLDER
			ENUMERATEDOCUMENTS Subfolder, sPath, sVpath
			
			' LIST FILES IN THE CURRENT FOLDER
			For each File in Subfolder.Files
				sHyperlink = sVirtualPath & "/" & Subfolder.Name & "/" & File.Name
				response.write  "<LI><img src=""menu/" & GetFileIcon( File.Name ) & """><A TARGET=""_NEW"" HREF=""" & sHyperlink  & """ class=""documentlist""> " & File.Name & " ( " & FormatNumber(File.Size,0) & " KB) </A></LI>"  & vbcrlf
			Next 

			' IF FOLDER CONTAINS NO FILES OR FOLDERS SHOW AS EMPTY
			If (Subfolder.Files.Count < 1) AND (Subfolder.SubFolders.Count < 1) Then
				response.write  "<LI> (<I> EMPTY </I>) </LI>"  & vbcrlf
			End IF


			' CLOSE FOLDER LIST TAG
			response.write "</UL>"

		Next


End Sub


'-------------------------------------------------------------------------------------------------------
' FUNCTION GETFILEICON(SNAME)
'-------------------------------------------------------------------------------------------------------
Function GetFileIcon(sName)

	sReturnValue = "images/msie.gif"

	Select Case lcase(right(sName,3))
		Case "doc"
            sReturnValue = "images/msword.gif"
        Case "xls"
            sReturnValue = "images/msexcel.gif"
        Case "ppt"
            sReturnValue = "images/msppt.gif"
        Case "htm"
            sReturnValue = "images/msie.gif"
        Case "pdf"
            sReturnValue = "images/pdf.gif"
	End Select

	  GetFileIcon = sReturnValue

End Function
%>
