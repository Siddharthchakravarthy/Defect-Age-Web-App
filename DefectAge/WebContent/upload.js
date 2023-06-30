function submitForm() {
  var form = document.getElementById("uploadForm");
  var formData = new FormData(form);

  var xhr = new XMLHttpRequest();
  xhr.open("POST", "/DefectAge/upload", true);
  xhr.responseType = "blob"; // Set the response type to blob
  xhr.onload = function () {
    if (xhr.status === 200) {
      // File upload and processing completed successfully
      var blob = xhr.response;
      var fileName = extractFileNameFromResponse(xhr.getResponseHeader("Content-Disposition"));

      // Create a temporary anchor element to trigger the file download
      var downloadLink = document.createElement("a");
      downloadLink.href = window.URL.createObjectURL(blob);
      downloadLink.download = fileName;
      downloadLink.click();

      // Clean up the temporary anchor element
      window.URL.revokeObjectURL(downloadLink.href);

      var responseMessage = "File download completed successfully";
      alert(responseMessage);
    }
  };
  xhr.send(formData);
}

// Helper function to extract the file name from the Content-Disposition header
function extractFileNameFromResponse(header) {
  var startIndex = header.indexOf("filename=\"") + 10;
  var endIndex = header.lastIndexOf("\"");
  var fileName = header.substring(startIndex, endIndex);
  return fileName;
}
