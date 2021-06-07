<!-- #include file="../../includes/common.asp" //-->
<!--#include file="../../includes/JSON_2.0.2.asp" //-->
<!-- #include file="../../include_top_functions.asp" //-->
<%

Dim a, sSql, sPath, iCounter

'sPath = "Meetings/Council/Minutes/2008/"
If request("path") = "" Then
	response.End 
End If 
sPath = request("path") & "/"
iCounter = clng(0)

sLocationName =  GetVirtualDirectyName()
'response.write "sLocationName = "  & sLocationName & "<br /><br />"

Set a = jsArray()

Set FSO = CreateObject("Scripting.FileSystemObject")

If FSO.FolderExists(Server.Mappath("/public_documents300/" & sLocationName & "/published_documents/" & sPath)) Then 
	Set oFolder = FSO.GetFolder(Server.Mappath("/public_documents300/" & sLocationName & "/published_documents/" & sPath))

	ShowFiles a, oFolder, Server.Mappath("/public_documents300/" & sLocationName & "/published_documents/" & sPath), "public_documents300/" & sLocationName & "/published_documents/" & sPath, iCounter

	If iCounter = clng(0) Then 
		' See if there are sub folders
		If FolderHasSubFolders( oFolder, Server.Mappath("/public_documents300/" & sLocationName & "/published_documents/" & sPath), "public_documents300/" & sLocationName & "/published_documents" & sPath ) Then 
			Set a(Null) = jsArray()
			a(Null)(Null) = "NOFILE"
			a(Null)(Null) = "NOFILE"
		Else 
			Set a(Null) = jsArray()
			a(Null)(Null) = "EMPTY"
			a(Null)(Null) = "EMPTY"
		End If 
	End If 
Else
	Set a(Null) = jsArray()
	a(Null)(Null) = "NOFILE"
	a(Null)(Null) = "NOFILE"
End If 

a.Flush

Set oFolder = Nothing 
Set FSO = Nothing 


'-------------------------------------------------------------------------------------------------------
' Sub ShowFiles( Folder, sPath, sVpath )
'-------------------------------------------------------------------------------------------------------
Sub ShowFiles( ByRef a, ByRef Folder, sPath, sVpath, ByRef iCounter )

	' BUILD HYPERLINK BASE PATH
	'sVirtualpath = Replace(Folder.Path,sPath,"")
	'sTempPath =  Replace(Folder.Path,sPath,"")
	'sVirtualpath = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & sVpath & replace(sVirtualpath,"\","/")
	sVirtualpath = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & sVpath

	' LIST CONTENTS OF FOLDER (FILES)
	For Each File In Folder.Files
		iCounter = iCounter + clng(1)
		' GET FILE SIZE
		If File.Size > 1024 Then
			sFileSize = FormatNumber((File.Size / 1024),0)  & " KB"
		Else
			sFileSize =  FormatNumber(File.Size,0) & " Bytes"
		End If

		'sHyperlink = sVirtualPath & "/" & Folder.Name & "/" & File.Name
		sHyperlink = sVirtualPath & File.Name
		'response.write vbcrlf & "<a class=""documentlist"" TARGET=""DOCUMENTS"" href=""" & sHyperlink  & """ > " & UCASE(File.Name) & " ( " & sFileSize & ") </a><br />"
		Set a(Null) = jsArray()
		a(Null)(Null) = sHyperlink
		a(Null)(Null) = UCASE(File.Name) & " (" & sFileSize & ")"
	Next

End Sub


'-------------------------------------------------------------------------------------------------------
' Function FolderHasSubFolders( oFolder )
'-------------------------------------------------------------------------------------------------------
Function FolderHasSubFolders( ByRef oFolder, sPath, sVpath )
	Dim SubFolder, iFolderCount

	iFolderCount = clng(0)

	For Each SubFolder In oFolder.SubFolders
		iFolderCount = iFolderCount + clng(1)
		Exit For 
	Next

	If iFolderCount > clng(0) Then 
		FolderHasSubFolders = True 
	Else
		FolderHasSubFolders = False 
	End If 

End Function 




%>
