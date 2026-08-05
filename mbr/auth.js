// MBR Dashboard Authentication Guard
// Include this script at the top of every MBR report page.
// Checks sessionStorage for auth and redirects to mbr/index.html if not logged in.
(function() {
  if (sessionStorage.getItem("mbr_auth") !== "1") {
    var path = location.pathname;
    // Month reports live at mbr/<Mon>/file.html and need "../index.html"; the
    // standalone analyses sit directly in mbr/ and need "./index.html". Match the
    // month segment generically -- the old hardcoded Jan/Feb/Mar/Apr list silently
    // sent May and June to a nonexistent mbr/<Mon>/index.html.
    // Folder names mix abbreviations and full names (Apr, May, June, March), so
    // match a month prefix plus whatever letters follow.
    var inMonthFolder = /\/(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\//.test(path);
    var redirect = inMonthFolder ? "../index.html" : "./index.html";
    window.location.replace(redirect);
  }
})();
