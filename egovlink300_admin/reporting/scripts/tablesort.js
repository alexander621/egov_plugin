/*

        TableSort by frequency-decoder.com

        Released under a creative commons nc-sa license (http://creativecommons.org/licenses/by-nc-sa/2.5/)

        Please credit frequency decoder in any derivitive work. Thanks.

        Changes
        -------

        14/02/2006 : Added the 'sort onload' functionality.
                     Integrated John Resigs addEvent function (http://ejohn.org/projects/flexible-javascript-events/).
        15/02/2006 : Added the ability to 'force' the script to use a specified sort algorithm.
                     Added a check for identical cell data (no sort is ran should all of a columns cells contain identical data).
        08/03/2006 : Added the ability to specify which column the table is initially sorted on.
        15/03/2006 : Added the className "sort-active" to the TH node and changed the cursor to "wait" during the sort process.
                     Integrated the addevent function into the fdTableSort object
        16/03/2006 : Added the ability to define custom sort functions within the TH nodes class attribute.
*/

var fdTableSort = {

        regExp_Currency:        /^[£$€]/,
        regExp_Number:          /^(\-)?[0-9]+(\.[0-9]*)?$/,
        pos:                    -1,

        addEvent: function(obj, type, fn) {
                if( obj.attachEvent ) {
                        obj["e"+type+fn] = fn;
                        obj[type+fn] = function(){obj["e"+type+fn]( window.event );}
                        obj.attachEvent( "on"+type, obj[type+fn] );
                } else
                        obj.addEventListener( type, fn, false );
        },
        init: function() {
                if (!document.getElementsByTagName) return;

                var tables = document.getElementsByTagName('table');
                var sortable, headers, columnNum = 0;

                for(var t = 0, tbl; tbl = tables[t]; t++) {
                        headers = tbl.getElementsByTagName('th');
                        sortable = false;
                        if(tbl.className.search(/sortable-onload-([0-9]+)/) != -1) {
                                columnNum = parseInt(tbl.className.match(/sortable-onload-([0-9]+)/)[1]) - 1;
                        }

                        for (var z=0, th; th = headers[z]; z++) {
                                if(th.className.match('sortable')) {
                                        if(tbl.className.match('sortable-onload') && z == columnNum) sortable = th;
                                        th.onclick = fdTableSort.sortWrapper;
                                        th.appendChild(document.createElement('span'));
                                }
                        }

                        if(sortable) fdTableSort.initSort(sortable);
                }
        },

        sortWrapper: function(e) {
                fdTableSort.initSort(this);
        },

        initSort: function(thNode) {

                var curr = thNode;

                thNode.className = thNode.className + " sort-active";
                document.getElementsByTagName('body')[0].style.cursor = "wait";

                fdTableSort.pos = 0;

                // Get the column position
                while(curr.previousSibling) {
                        if(curr.previousSibling.nodeType != 3) fdTableSort.pos++;
                        curr = curr.previousSibling;
                }

                // Remove any old "reverse" class we might have previously added
                var thCollection = curr.parentNode.getElementsByTagName('th');

                var span;

                for(var i=0, th; th = thCollection[i]; i++) {
                        if(i != fdTableSort.pos) {
                                th.className = th.className.replace('reverseSort','');

                                // Remove arrow
                                span = th.getElementsByTagName('span');

                                if(span.length > 0) {
                                        span = span[span.length - 1];
                                        if(span.firstChild) span.removeChild(span.firstChild);
                                        span.appendChild(document.createTextNode("\u00a0\u00a0"));
                                }
                        }
                }

                // Get the table
                var tableElem = thNode;
                while(tableElem.tagName.toLowerCase() != 'table' && tableElem.parentNode) {
                        tableElem = tableElem.parentNode;
                }

                // Has a row color been defined in the table's className?
                var style;

                if(tableElem.className.match(/style-(.+)/)) {
                        style = tableElem.className.match(/style-(.+)/)[1];
                }

                // Has the table a tbody ?
                // N.B. By definition, tables can have multiple tbody tags
                // this script assumes only one.
                if(tableElem.getElementsByTagName('tbody')) {
                        tableElem = tableElem.getElementsByTagName('tbody')[0];
                }

                // Get the tr tags
                var trs = tableElem.getElementsByTagName('tr');
                var trCollection = new Array();

                // If the current tr has any th child elements then skip it..
                for(var i = 0, tr; tr = trs[i]; i++) {
                        if(tr.getElementsByTagName('th').length == 0) trCollection.push(tr);
                }

                // Try to get the column data type
                var sortFunction;
                var txt         = null;
                var identical   = true;
                var firstTxt    = "";

                for(i = 0; i < trCollection.length; i++) {
                        cellTxt = fdTableSort.getInnerText(trCollection[i].getElementsByTagName('td')[fdTableSort.pos]);
                        if(i > 0 && txt != cellTxt) identical = false;
                        txt = cellTxt;
                        if(firstTxt == "") firstTxt = txt;
                }

                if(thNode.className.match('sortable-numeric'))          sortFunction = fdTableSort.sortNumeric;
                else if(thNode.className.match('sortable-currency'))    sortFunction = fdTableSort.sortCurrency;
                else if(thNode.className.match('sortable-date'))        sortFunction = fdTableSort.sortDate;
                else if(thNode.className.search(/sortable-([a-zA-Z\_]+)/) != -1 && thNode.className.match(/sortable-([a-zA-Z\_]+)/)[1] in window) sortFunction = window[thNode.className.match(/sortable-([a-zA-Z\_]+)/)[1]];
                else if(fdTableSort.dateFormat(firstTxt) != 0)          sortFunction = fdTableSort.sortDate;
                else if(firstTxt.match(fdTableSort.regExp_Number))      sortFunction = fdTableSort.sortNumeric;
                else if(firstTxt.match(fdTableSort.regExp_Currency))    sortFunction = fdTableSort.sortCurrency;
                else                                                    sortFunction = fdTableSort.sortCaseInsensitive;

                // Call the JavaScript array.sort method, passing in our bespoke sort function
                if(!identical) trCollection.sort(sortFunction);

                // Do we need to reverse the sort?
                var arrow;

                if(thNode.className.match('reverseSort') && !identical) {
                        trCollection.reverse();
                        thNode.className = thNode.className.replace('reverseSort','');
                        arrow = " \u2191";
                } else {
                        thNode.className = thNode.className.replace('reverseSort','');
                        thNode.className = thNode.className + ' reverseSort';
                        arrow = " \u2193";
                }

                span = thNode.getElementsByTagName('span');

                if(span.length > 0) {
                        span = span[span.length - 1];
                        if(span.firstChild) span.removeChild(span.firstChild);
                        span.appendChild(document.createTextNode(arrow));
                }

                document.getElementsByTagName('body')[0].style.cursor = "auto";
                thNode.className = thNode.className.replace("sort-active", "");

                if(identical) return;

                for(var i = 0, tr; tr = trCollection[i]; i++) {
                        if(style) {
                                tr.className = tr.className.replace(style,'');
                                tr.className = (i % 2 != 0) ? tr.className + " " + style : tr.className;
                        }
                        tableElem.appendChild(tr);
                }
        },

        getInnerText: function(el) {
                if (typeof el == "string" || typeof el == "undefined") return el;
                if(el.innerText) return el.innerText;
                var txt = '', i;
                for (i = el.firstChild; i; i = i.nextSibling) {
                        if (i.nodeType == 3)            txt += i.nodeValue;
                        else if (i.nodeType == 1)       txt += fdTableSort.getInnerText(i);
                }
                return txt;
        },

        dateFormat: function(dateIn) {

                var y, m, d, res;

                if(dateIn.match(/^(0[1-9]|1[012])([- \/.])(0[1-9]|[12][0-9]|3[01])([- \/.])(\d\d?\d\d)$/)) {
                        res = dateIn.match(/^(0[1-9]|1[012])([- \/.])(0[1-9]|[12][0-9]|3[01])([- \/.])(\d\d?\d\d)$/);
                        y = res[5];
                        m = res[1];
                        d = res[3];
                } else if(dateIn.match(/^(0[1-9]|[12][0-9]|3[01])([- \/.])(0[1-9]|1[012])([- \/.])(\d\d?\d\d)$/)) {
                        res = dateIn.match(/^(0[1-9]|[12][0-9]|3[01])([- \/.])(0[1-9]|1[012])([- \/.])(\d\d?\d\d)$/);
                        y = res[5];
                        m = res[3];
                        d = res[1];
                } else if(dateIn.match(/^(\d\d?\d\d)([- \/.])(0[1-9]|1[012])([- \/.])(0[1-9]|[12][0-9]|3[01])$/)) {
                        res = dateIn.match(/^(\d\d?\d\d)([- \/.])(0[1-9]|1[012])([- \/.])(0[1-9]|[12][0-9]|3[01])$/);
                        y = res[1];
                        m = res[3];
                        d = res[5];
                } else return 0;

                if(m.length == 1) m = "0" + m;
                if(d.length == 1) d = "0" + d;
                if(y.length != 4) y = (parseInt(y) < 50) ? '20' + y : '19' + y;

                return y+m+d;
        },

        sortDate: function(a,b) {
                aa = fdTableSort.dateFormat(fdTableSort.getInnerText(a.getElementsByTagName('td')[fdTableSort.pos]));
                bb = fdTableSort.dateFormat(fdTableSort.getInnerText(b.getElementsByTagName('td')[fdTableSort.pos]));

                // Added 15/02/2006 : If mixed type row then treat the date as being bigger
                if(aa == 0 && bb != 0) return -1;
                else if(bb == 0 && aa != 0) return 1;

                if (aa == bb) return 0;
                if (aa < bb)  return -1;
                return 1;
        },

        sortCurrency:function(a,b) {
                aa = fdTableSort.getInnerText(a.getElementsByTagName('td')[fdTableSort.pos]).replace(/[^0-9.]/g,'');
                bb = fdTableSort.getInnerText(b.getElementsByTagName('td')[fdTableSort.pos]).replace(/[^0-9.]/g,'');

                // Added 15/02/2006 : If mixed type row then treat the number as being bigger
                if((isNaN(aa) || aa == "") && !isNaN(bb)) return -1;
                else if((isNaN(bb) || bb == "") && !isNaN(aa)) return 1;

                if(isNaN(aa) || aa == "") aa = 0;
                if(isNaN(bb) || bb == "") bb = 0;

                return parseFloat(aa) - parseFloat(bb);
        },

        sortNumeric:function (a,b) {
                aa = parseFloat(fdTableSort.getInnerText(a.getElementsByTagName('td')[fdTableSort.pos]));
                bb = parseFloat(fdTableSort.getInnerText(b.getElementsByTagName('td')[fdTableSort.pos]));

                // Added 15/02/2006 : If mixed type row then treat the number as being bigger
                if((isNaN(aa) || aa == "") && !isNaN(bb)) return -1;
                else if((isNaN(bb) || bb == "") && !isNaN(aa)) return 1;

                if(isNaN(aa) || aa == "") aa = 0;
                if(isNaN(bb) || bb == "") bb = 0;

                return aa-bb;
        },

        sortCaseInsensitive:function (a,b) {
                aa = fdTableSort.getInnerText(a.getElementsByTagName('td')[fdTableSort.pos]).toLowerCase();
                bb = fdTableSort.getInnerText(b.getElementsByTagName('td')[fdTableSort.pos]).toLowerCase();
                if(aa == bb) return 0;
                if(aa < bb)  return -1;
                return 1;
        }
}

fdTableSort.addEvent(window, "load", fdTableSort.init);







