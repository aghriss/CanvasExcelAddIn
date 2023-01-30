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
    /*
      const url = "http://localhost:7071/api/AddTwo"/!**!/;
    */

    /*  return new Promise(async function (resolve, reject) {
        try {
          //Note that POST uses text/plain because custom functions runtime does not support full CORS
          const response = await fetch(url, {
            method: "POST",
            headers: {
              "Content-Type": "text/plain",
            },
            body: JSON.stringify({ first: first, second: second }),
          });
          const jsonAnswer = await response.json();
          resolve(jsonAnswer.answer);
        } catch (error) {
          console.log("error", error.message);
        }
      });*/
    return first + second
}

/**
 *
 * @param token {string}
 * @param endpoint {string}
 * @returns {json}
 */
function canvasAPI(token, endpoint) {
   var url = "https://canvas.colorado.edu/api/v1/" + endpoint;
   //Logger.log(url);
   //Logger.log(token);
   var params = {
     "muteHttpExceptions": true,
      "headers": {
         "Authorization": "Bearer " + token,
      },
      "method": "GET"
   }
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
function canvasTokenUser(token){
  let endpoint = "users/self";
  let response = canvasAPI(token, endpoint);
  //Logger.log(response);
  return response["name"]
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
function canvasGrade(token, course_id, assignment_id,student_id){
  let grade = parseFloat(canvasAPI(token, "courses/" +course_id + "/assignments/" + assignment_id + "/submissions/" + student_id)["grade"]);
  if(isNaN(grade)) return 0;
  return grade
}

function test(){
  let token = "10772~nd3oDky72Zu4YxfB2PTfYfBSLcaWfyq3GaVMrnyXOwBLERIDfLg3REEMkLUrPrSZ"
  Logger.log(canvasTokenUser(token));
  Logger.log(canvasGrade(313886,1455469, 83599,token));
}

CustomFunctions.associate("ADD", add);
CustomFunctions.associate("Grade", canvasGrade);
CustomFunctions.associate("GraderName", canvasTokenUser);
CustomFunctions.associate("ADD", add);
