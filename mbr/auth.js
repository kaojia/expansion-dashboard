// MBR Dashboard Authentication Guard
// Include this script at the top of every MBR report page.
// Checks sessionStorage for auth and redirects to mbr/index.html if not logged in.
(function() {
  if (sessionStorage.getItem("mbr_auth") !== "1") {
    var path = location.pathname;
    // HTML files are at mbr/Apr/file.html -> go up one level to mbr/index.html
    var redirect = "./index.html";
    if (path.indexOf("/Apr") !== -1 || path.indexOf("/Mar") !== -1 || path.indexOf("/Feb") !== -1 || path.indexOf("/Jan") !== -1) {
      redirect = "../index.html";
    }
    window.location.replace(redirect);
  }
})();
