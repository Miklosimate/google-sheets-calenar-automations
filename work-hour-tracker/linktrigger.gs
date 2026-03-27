function doGet(e) {
  try {
    createWeeklyProgressEvent(); // Call your function
    return ContentService.createTextOutput("Success! Weekly progress event updated.");
  } catch (error) {
    return ContentService.createTextOutput("Error: " + error.message);
  }
}

