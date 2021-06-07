function OpenFolder(iIndex,iParentFolder){

// IF NOT IE CREATE ARRAY OF UL ELEMENTS
if (ns6){
for (i=0;i<document.getElementsByTagName("UL").length;i++){
if (document.getElementsByTagName("UL")[i].id=="foldinglist"){
foldercontentarray[c]=document.getElementsByTagName("UL")[i]
c++
}
}
}

// OPEN SELECTED FOLDER
if (iIndex >= 0)
{
	if (ns6){
		foldercontentarray[iIndex].style.display=''
		foldercontentarray[iIndex].previousSibling.previousSibling.style.listStyleImage="url(../open.gif)"
	}
	else{

		if (iParentFolder > - 1)
		{
		for(i=0;i<(iIndex-iParentFolder)+1;i++){
		foldinglist[iParentFolder+i].style.display=''
		document.all[foldinglist[iParentFolder+i].sourceIndex -1].style.listStyleImage="url(../open.gif)"}
		//OPEN PARENT
		//foldinglist[iParentFolder+i].style.display=''
		//document.all[foldinglist[iParentFolder+i].sourceIndex -1].style.listStyleImage="url(../open.gif)

		//foldinglist[iParentFolder+i].style.display=''
		//document.all[foldinglist[iParentFolder+i].sourceIndex -1].style.listStyleImage="url(../open.gif);
		
		
		
		}
		else
		{
		foldinglist[iIndex].style.display=''
		document.all[foldinglist[iIndex].sourceIndex -1].style.listStyleImage="url(../open.gif)"}
	
}
}

}


