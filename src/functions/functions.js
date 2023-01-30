// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

/**
 * Canvas API base
 * costumfunction
 * @param token {string}
 * @param endpoint {string}
 * @returns {json}
 */
function canvasAPI(token, endpoint) {
  var url = "https://canvas.colorado.edu/api/v1/" + endpoint;
  //Logger.log(url);
  //Logger.log(token);
  var params = {
    muteHttpExceptions: true,
    headers: {
      Authorization: "Bearer " + token,
    },
    method: "GET",
  };
  //Logger.log(params);
  var response = UrlFetchApp.fetch(url, params);
  //Logger.log(response);
  var json = response.getContentText();
  return JSON.parse(json);
}

/**
 https://officedev.github.io/Office-Add-in-samples/Excel-custom-functions/AzureFunction/CustomFunctionProject/f

 * @param token {string}
 * @returns {string}
 */
function canvasTokenUser(token) {
  let endpoint = "users/self";
  let response = canvasAPI(token, endpoint);
  //Logger.log(response);
  return response["name"];
}

/**
 * Get Grade
 * @customfunction
 * @param token {string}
 * @param course_id {number}
 * @param assignment_id {number}
 * @param student_id {number}
 * @returns {number}
 */
function canvasGrade(token, course_id, assignment_id, student_id) {
  let grade = parseFloat(
    canvasAPI(token, "courses/" + course_id + "/assignments/" + assignment_id + "/submissions/" + student_id)["grade"]
  );
  if (isNaN(grade)) return 0;
  return grade;
}

CustomFunctions.associate("ADD", add);
CustomFunctions.associate("Grade", canvasGrade);
CustomFunctions.associate("Name", canvasTokenUser);
