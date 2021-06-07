function addFieldRow(iIsRootAdmin, iIsLimited, iTableID, iTotalFieldsID, iRowID, iPageType) {
  var mytbl           = document.getElementById(iTableID);
  var totalrows       = Number(document.getElementById(iTotalFieldsID).value);
  var lcl_isRootAdmin = '';
  var lcl_isLimited   = '';
  var lcl_pageType    = '';   //pagetypes: SECTION, MPT (MapPoint Type)

  //Determine if this a root admin
  if((iIsRootAdmin == "")||(iIsRootAdmin == undefined)) {
      lcl_isRootAdmin = 'False';
  } else {
      lcl_isRootAdmin = iIsRootAdmin;
  }

  //Determine if this is a limited view
  if((iIsLimited == "")||(iIsLimited == undefined)) {
      lcl_isLimited = 'False';
  } else {
      lcl_isLimited = iIsLimited;
  }

  //Determine which page type to set up for.
  if((iPageType == "")||(iPageType == undefined)) {
      lcl_pageType = 'SECTION';
  } else {
      lcl_pageType = iPageType;
  }

  //Increase the total rows by one.  This is index for the new row.
  totalrows = totalrows+1;

  //Set up the new row.
  var row = mytbl.insertRow(totalrows);
      row.id = iRowID + totalrows;

  //Set the background color.  Odd rows: "#eeeeee", Even rows: "#ffffff"
  var lcl_rowbg   = "";
  var lcl_evenodd = totalrows/2;
      lcl_evenodd = lcl_evenodd.toString();

  if(lcl_evenodd.indexOf('.') > 0) {
     lcl_rowbg = "#eeeeee";
  }else{
     lcl_rowbg = "#ffffff";
  }

  row.style.background = lcl_rowbg;

  //BEGIN: Build the cells for the new row. -----------------------------------
  if(lcl_pageType == 'MPT') {
     var a = row.insertCell(0);  //Row Count
     var b = row.insertCell(1);  //Field Name
     var c = row.insertCell(2);  //Display In Results
     var d = row.insertCell(3);  //Display On Info Page
     var e = row.insertCell(4);  //Results Order
     var f = row.insertCell(5);  //In Public Search
     var g = row.insertCell(6);  //Delete Row and Additional Info
  } else {
     var a = row.insertCell(0);  //Row Count
     var b = row.insertCell(1);  //Field Name
     var c = row.insertCell(2);  //Active
     var d = row.insertCell(3);  //Display Order
     var e = row.insertCell(4);  //Display Field Name (Display Label)
     var f = row.insertCell(5);  //Display as Multi-Line
     var g = row.insertCell(6);  //Include "Add a Link"
     var h = row.insertCell(7);  //Delete Row and Additional Info
  }
  //END: Build the cells for the new row. -------------------------------------

  //BEGIN: Build cells --------------------------------------------------------

  //Delete Row and Additional Info
  var cell_deleterow = document.createElement('input');
      cell_deleterow.type      = 'checkbox';
      cell_deleterow.name      = 'deleteField' + totalrows;
      cell_deleterow.id        = 'deleteField' + totalrows;
      cell_deleterow.value     = 'Y';
      cell_deleterow.checked   = '';

  //Build page specific columns -----------------------------------------------
  if(lcl_pageType == 'MPT') {

     //Display In Results
     var cell_displayresults = document.createElement('input');
         cell_displayresults.type      = 'checkbox';
         cell_displayresults.name      = 'displayInResults' + totalrows;
         cell_displayresults.id        = 'displayInResults' + totalrows;
         cell_displayresults.value     = '1';
         cell_displayresults.checked   = 'checked';

     //Display On Info Page
     var cell_displayonpage = document.createElement('input');
         cell_displayonpage.type      = 'checkbox';
         cell_displayonpage.name      = 'displayInInfoPage' + totalrows;
         cell_displayonpage.id        = 'displayInInfoPage' + totalrows;
         cell_displayonpage.value     = '1';
         cell_displayonpage.checked   = 'checked';

     //Results Order
     var cell_resultsorder = document.createElement('input');
         cell_resultsorder.type      = 'text';
         cell_resultsorder.name      = 'resultsOrder' + totalrows;
         cell_resultsorder.id        = 'resultsOrder' + totalrows;
         cell_resultsorder.size      = '3';
         cell_resultsorder.maxLength = '5';
         cell_resultsorder.value     = totalrows;

     //In Public Search
     var cell_publicsearch = document.createElement('input');
         cell_publicsearch.type      = 'checkbox';
         cell_publicsearch.name      = 'inPublicSearch' + totalrows;
         cell_publicsearch.id        = 'inPublicSearch' + totalrows;
         cell_publicsearch.value     = '1';
         cell_publicsearch.checked   = '';

     var cell_deleterow2 = document.createElement('input');
         cell_deleterow2.type      = 'hidden';
         cell_deleterow2.name      = 'mp_fieldid' + totalrows;
         cell_deleterow2.id        = 'mp_fieldid' + totalrows;
         cell_deleterow2.size      = '5';
         cell_deleterow2.maxLength = '';

  } else {

     //Field Name
     var cell_fieldname = document.createElement('input');
         cell_fieldname.type      = 'text';
         cell_fieldname.name      = 'fieldname' + totalrows;
         cell_fieldname.id        = 'fieldname' + totalrows;
         cell_fieldname.size      = '50';
         cell_fieldname.maxLength = '100';
         cell_fieldname.onchange  = function() { clearMsg('fieldname' + totalrows); };

     if(lcl_isRootAdmin == 'True') {
        var cell_fieldtype = document.createElement('input');
            cell_fieldtype.type      = 'text';
            cell_fieldtype.name      = 'fieldtype' + totalrows;
            cell_fieldtype.id        = 'fieldtype' + totalrows;
            cell_fieldtype.size      = '15';
            cell_fieldtype.maxLength = '100';
            cell_fieldtype.onchange  = function() { clearMsg('fieldtype' + totalrows); };

        var cell_fieldtype_label1 = document.createElement('span');
            cell_fieldtype_label1.innerHTML = '<br /><strong>Field Type: </strong>(code use ONLY)&nbsp;';
     }

     //Active
     var cell_isActive = document.createElement('input');
         cell_isActive.type      = 'checkbox';
         cell_isActive.name      = 'sectionfield_isActive' + totalrows;
         cell_isActive.id        = 'sectionfield_isActive' + totalrows;
         cell_isActive.value     = 'Y';
         cell_isActive.checked   = 'checked';

     //Display Order
     var cell_displayorder = document.createElement('input');
         cell_displayorder.type      = 'text';
         cell_displayorder.name      = 'displayOrder' + totalrows;
         cell_displayorder.id        = 'displayOrder' + totalrows;
         cell_displayorder.size      = '3';
         cell_displayorder.maxLength = '5';
         cell_displayorder.value     = totalrows;

     //Display Field Name (Display Label)
     var cell_displayFieldName = document.createElement('input');
         cell_displayFieldName.type      = 'checkbox';
         cell_displayFieldName.name      = 'displayFieldName' + totalrows;
         cell_displayFieldName.id        = 'displayFieldName' + totalrows;
         cell_displayFieldName.value     = 'Y';
         cell_displayFieldName.checked   = 'checked';

     //Display as Multi-Line
     var cell_multiline = document.createElement('input');
         cell_multiline.type      = 'checkbox';
         cell_multiline.name      = 'isMultiLine' + totalrows;
         cell_multiline.id        = 'isMultiLine' + totalrows;
         cell_multiline.value     = '1';
         cell_multiline.checked   = '';

     //Include "Add a Link"
     var cell_addalink = document.createElement('input');
         cell_addalink.type      = 'checkbox';
         cell_addalink.name      = 'hasAddLinkButton' + totalrows;
         cell_addalink.id        = 'hasAddLinkButton' + totalrows;
         cell_addalink.value     = '1';
         cell_addalink.checked   = '';

     var cell_deleterow2 = document.createElement('input');
         cell_deleterow2.type      = 'hidden';
         cell_deleterow2.name      = 'section_fieldid' + totalrows;
         cell_deleterow2.id        = 'section_fieldid' + totalrows;
         cell_deleterow2.size      = '5';
         cell_deleterow2.maxLength = '';
  }
  //END: Build cells ----------------------------------------------------------

  //BEGIN: Display the cells to the row ---------------------------------------

  //Row Count
  a.innerHTML = totalrows + '. ';

  if(lcl_pageType == 'MPT') {

     c.align = 'center';
     d.align = 'center';
     e.align = 'center';
     f.align = 'center';
     g.align = 'center';

     b.innerHTML = '&nbsp;';
     c.appendChild(cell_displayresults);
     d.appendChild(cell_displayonpage);
     e.appendChild(cell_resultsorder);
     f.appendChild(cell_publicsearch);
     g.appendChild(cell_deleterow);
     g.appendChild(cell_deleterow2);

  } else {

     c.align = 'center';
     d.align = 'center';
     e.align = 'center';
     f.align = 'center';
     g.align = 'center';
     h.align = 'center';

     b.appendChild(cell_fieldname);
     c.appendChild(cell_isActive);
     d.appendChild(cell_displayorder);
     e.appendChild(cell_displayFieldName);
     f.appendChild(cell_multiline);
     g.appendChild(cell_addalink);
     h.appendChild(cell_deleterow);
     h.appendChild(cell_deleterow2);


     if(lcl_isRootAdmin == 'True') {
        b.align='right';
        b.appendChild(cell_fieldtype_label1);
        b.appendChild(cell_fieldtype);
     }
  }
  //END: Display the cells to the row -----------------------------------------

  //update the total row count.
  document.getElementById(iTotalFieldsID).value = totalrows;
}
