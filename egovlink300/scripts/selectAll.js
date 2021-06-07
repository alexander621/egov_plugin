//function selectAll(frm, value [,sString])
//The optional value sString is the text to look for in the form element name
//Used to select or deselect all checkboxes that contain del_ in their name
function selectAll(frm, value)
  {
   var searchString="del_";
   if(selectAll.arguments.length == 3)searchString=selectAll.arguments[2];
   var formObject=eval("document." + frm);
   var arrElms=formObject.all;
   for(var i=0; i<arrElms.length; i++)
   {
     if(arrElms[i].name && arrElms[i].type=="checkbox" && arrElms[i].name.indexOf(searchString) >= 0)
       arrElms[i].checked=value;
   }
  }