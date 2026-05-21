$(function () {
  $("#header").load("includes/header.html");
  $("#footer").load("includes/footer.html");
});


var currentLat = "";
var currentLng = "";

function getMyLocation() {
  if (!navigator.geolocation) {
    alert("Geolocation is not supported by this browser.");
    return;
  }
  navigator.geolocation.getCurrentPosition(
    function (pos) {
      currentLat = pos.coords.latitude.toFixed(6);
      currentLng = pos.coords.longitude.toFixed(6);
      var msg = "Location captured:\nLat: " + currentLat + "\nLng: " + currentLng;
      msg += "\n\nClick \"Save Place\" to store this location.";
      alert(msg);
    },
    function () {
      alert("Unable to retrieve your location.");
    }
  );
}

function saveData() {
  var color    = $("#selColor").val();
  var section  = $("#txtSection").val().trim();
  var number   = $("#txtNumber").val().trim();

  if (!section || !number || color === "Selecione") {
    alert("Please fill in all fields. Location is optional.");
    return;
  }

  var now = new Date().toLocaleString();

  localStorage.setItem("Park_Date",    now);
  localStorage.setItem("Park_Color",   color);
  localStorage.setItem("Park_Section", section);
  localStorage.setItem("Park_Number",  number);
  localStorage.setItem("Park_Lat",     currentLat);
  localStorage.setItem("Park_Lng",     currentLng);

  $("#txtSection").val("");
  $("#txtNumber").val("");
  $("#selColor").val("Selecione");
  currentLat = "";
  currentLng = "";

  $("#MessageModal").modal("show");
}

function showData() {
  var date    = localStorage.getItem("Park_Date")    || "N/A";
  var color   = localStorage.getItem("Park_Color")   || "N/A";
  var section = localStorage.getItem("Park_Section") || "N/A";
  var number  = localStorage.getItem("Park_Number")  || "N/A";
  var lat     = localStorage.getItem("Park_Lat")     || "";
  var lng     = localStorage.getItem("Park_Lng")     || "";

  $("#txtDate").text(date);
  $("#txtColor").text(color);
  $("#txtSection").text(section);
  $("#txtNumber").text(number);

  if (lat && lng) {
    $("#txtLat").text(lat);
    $("#txtLng").text(lng);
    $("#coordRow").show();
  } else {
    $("#coordRow").hide();
  }

  var bg, text;
  switch (color) {
    case "Yellow": bg = "#f9ca24"; text = "#212529"; break;
    case "Blue":   bg = "#3498db"; text = "#ffffff"; break;
    case "Orange": bg = "#e67e22"; text = "#ffffff"; break;
    case "Green":  bg = "#27ae60"; text = "#ffffff"; break;
    case "Red":    bg = "#e74c3c"; text = "#ffffff"; break;
    case "N/A":    bg = "#dee2e6"; text = "#212529"; break;
    default:       bg = "#dee2e6"; text = "#212529";
  }

  if (color !== "N/A") {
    $("#divData").css({
      "background-color": bg,
      "color": text,
      "border-radius": "8px",
      "padding": "20px"
    });
  }

  if (lat && lng) {
    setTimeout(function () {
      var mapEl = document.getElementById("parkMap");
      if (mapEl && !window.getMapAttrs) {
        window.parkMapSrc =
          "https://www.google.com/maps?q=" + lat + "," + lng +
          "&z=17&output=embed";
        mapEl.src = window.parkMapSrc;
      }
    }, 300);
  }
}
