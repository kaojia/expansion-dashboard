// NSR Dashboard Authentication Guard
// Checks sessionStorage for auth and redirects to nsr/index.html if not logged in.
(function() {
  if (sessionStorage.getItem("nsr_auth") !== "1") {
    var path = location.pathname;
    if (path.indexOf("/nsr/") !== -1) {
      window.location.replace("index.html");
    } else {
      window.location.replace("nsr/index.html");
    }
  }
})();
