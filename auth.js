// WBR Dashboard Authentication Guard
// Include this script at the top of every WBR report page.
// It checks sessionStorage for auth and redirects to index.html if not logged in.
(function() {
  if (sessionStorage.getItem("wbr_auth") !== "1") {
    // Not authenticated — redirect to login page
    // Determine the base path (works for W16/W16_WBR_AE_Pipeline.html -> ../index.html)
    var depth = (location.pathname.match(/\//g) || []).length;
    var base = "";
    // For GitHub Pages: /expansion-dashboard/W16/file.html
    // We need to go up to the W## folder level
    var path = location.pathname;
    var parts = path.split("/").filter(Boolean);
    // Find how many levels deep we are from the root index.html
    // The HTML files are at W16/file.html, so one level deep
    var redirect = "./index.html";
    if (path.indexOf("/W") !== -1) {
      redirect = "../index.html";
    }
    window.location.replace(redirect);
  }
})();
