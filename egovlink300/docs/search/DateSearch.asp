<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<P>&nbsp;</P>
<%
' SET VARIABLES
path = "..\pub"
strSearch = 1
session("strResults") = ""

' 1 = Last week files
' 2 = Today's files

Response.Write  DateSearch(path,2)

%>
</BODY>
</HTML>
<%
'********************FUNCTIONS*******************************************************************
Function DateSearch(path,intRange)
  
  'only tranlate to a local path if a virtual path is specified
  If InStr(1, path, ":") < 1 Then
    fullPath = Server.MapPath(path)
  Else
    fullPath = path
  End If

  Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

  blnFound = objFSO.FolderExists(fullPath)
  If blnFound Then
	  Set objDir = objFSO.GetFolder(fullPath)
  End If

 
  If IsObject(objDir) And Err.Number = 0 Then
    For Each objFound in objDir.SubFolders
      'Response.Write objFound
      For Each objFiles In objFound.Files
	    myDate = Now()
	    
	    Select Case intRange
	    Case 1
	        ' MATCH FILES ADDED FOR LAST WEEK
			dtTimeFrame= DateAdd("ww",-1,mydate)
	    Case 2
			' MATCH FILES ADDED TODAY
			dtTimeFrame = DateAdd("h",-24,mydate)
	    End Select 
	       
	    If objFiles.DateCreated >= dtTimeFrame Then
			session("strResults") = session("strResults") & "<li><a href='"&MapURL(objFiles.Path)&"'>" & objFiles.Name & "</a> created on " & objFiles.DateCreated & "</b>"
		End If
	  Next
      DateSearch objFound,intRange 
    Next
    Set objDir = Nothing
    Set objFSO = Nothing

  Else
    Set objDir = Nothing
    Set objFSO = Nothing
    Response.Write "Error Occured."
  End If
  
  If session("strResults") <> "" Then
		DateSearch = session("strResults")
  Else
		DateSearch = "Could not find any documents that meet that search criteria."
  End If 
 
  
  
End Function

Function MapURL(path)
     dim rootPath, url
     'Convert a physical file path to a URL for hypertext links.
     rootPath = Server.MapPath("/")
     url = Right(path, Len(path) - Len(rootPath))
     MapURL = Replace(url, "\", "/")
End Function 
%>