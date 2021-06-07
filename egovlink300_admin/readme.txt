If there is a file in the root directory with the extension .sql, you must run that in Query Analyzer against the boardsite database.

These are the bugs fixed in this release.

3.0.6 Fixes
Annotation files were missing in the release version of ecBoardSite.  These files were replaced.

The previous fix to the “Secure/non-secure” alert broke the documents section.  Now the “iframe” has a fully qualified path.

There were broken images in the delete folder page.  This image was added to the site.


3.0.5 Changes
Users can now add secure links as favorites.

Also, if you try to change the language, it will not break the site.

If you are running this site on a secure server, then you will nolonger get a warning about insecure items.


3.0.4 Fixes
Fixed a bug with stored procedures in MSSQL that prevented users of multiple groups to see all of the groups available.

3.0.3 Fixes
The pop up for editing the security of a directory now has a scrollbar for when you have more directories than the pop up allows.

3.0.2 Fixes
You can add a new document as an agenda item that will also become a document in the documents section.

Also, you are now able to edit the security of a poll created in the meetings section.



3.0.1 Fixes
Fixed the directory searching so that it doesn’t error out.

Also allowed searching to search additionally on other fields such as email, first name, zip code.

