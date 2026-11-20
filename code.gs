/********************************************
 * SMART AI LEAD MANAGEMENT SYSTEM
 * Google Forms + Google Sheets + Gemini AI
 * + Gmail + Google Apps Script
 ********************************************/

// 1. CONFIGURATION
const CONFIG = {
  MODEL: "gemini-3.5-flash-lite",
  RESPONSE_COLUMN: "AI Generated Response",
  STATUS_COLUMN: "Automation Status"
};

// 2. MAIN AUTOMATION
// Runs automatically whenever someone submits the form.
function onFormSubmit(e) {

  if (!e || !e.range || !e.namedValues) {
    throw new Error(
      "This function must run from a form submission trigger."
    );
  }

  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const data = e.namedValues;

  // Read customer information
  const name = getAnswer(data, "Full Name");
  const email = getAnswer(data, "Email Address");
  const phone = getAnswer(data, "Phone Number");
  const service = getAnswer(data, "Service Interested In");
  const budget = getAnswer(data, "Budget");
  const requirements = getAnswer(data, "Your Requirements");

  // Prepare output columns
  const responseCol = getOrCreateColumn(
    sheet,
    CONFIG.RESPONSE_COLUMN
  );

  const statusCol = getOrCreateColumn(
    sheet,
    CONFIG.STATUS_COLUMN
  );

  try {

    // Validate customer details
    if (!name || !email) {
      throw new Error("Customer name or email is missing.");
    }

    // Generate personalised AI response
    const aiResponse = generateAIResponse({
      name: name,
      email: email,
      phone: phone,
      service: service,
      budget: budget,
      requirements: requirements
    });

    // Send personalised email
    GmailApp.sendEmail(
      email,
      "Thank you for your interest, " + name + "!",
      aiResponse,
      {
        name: "AI Lead Management Team"
      }
    );

    // Save AI response in Google Sheets
    sheet.getRange(row, responseCol)
      .setValue(aiResponse);

    // Update automation status
    sheet.getRange(row, statusCol)
      .setValue("Email Sent Successfully");

    Logger.log("Lead processed successfully: " + name);

  } catch (error) {

    // Record errors in Google Sheets
    sheet.getRange(row, statusCol)
      .setValue("Failed: " + error.message);

    Logger.log("Automation Error: " + error.message);
  }
}


// 3. GET FORM ANSWERS
function getAnswer(data, question) {

  // Match headings with or without question numbers
  const matchingKey = Object.keys(data).find(key =>
    key.replace(/^\d+\.\s*/, "").trim() === question.trim()
  );

  if (!matchingKey) {
    return "";
  }

  const answer = data[matchingKey];

  if (!answer || !answer.length) {
    return "";
  }

  return String(answer[0]).trim();
}


// 4. GENERATE AI RESPONSE USING GEMINI
function generateAIResponse(lead) {

  // 1. Get API key
  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty("GEMINI_API_KEY");

  if (!apiKey) {
    throw new Error("Gemini API key not found.");
  }

  // 2. Prepare AI prompt
  const prompt = `
You are a professional business sales assistant
representing Tecnosol Pvt Ltd.

Write a personalised and professional email
for the following potential customer.

Customer Name: ${lead.name}
Email: ${lead.email}
Phone: ${lead.phone}
Interested Service: ${lead.service}
Customer Budget: ${lead.budget}
Customer Requirements: ${lead.requirements}

Instructions:
1. Greet the customer by their name.
2. Thank them for their interest.
3. Mention only their selected service.
4. Acknowledge their requirements.
5. Refer to their budget naturally.
6. Invite them to discuss further.
7. Use a friendly professional tone.
8. Keep the email between 70 and 120 words.
9. Do not invent prices or promises.
10. Return only the email body.
11. End every email with this exact signature:

Best regards,
Yogesh
Tecnosol Pvt Ltd

12. Never write [Your Name] or [Your Company].
13. Never invent or change the company name.
`;

  // 3. Gemini API configuration
  const model = CONFIG.MODEL;

  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/"
    + model + ":generateContent";

  const payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }]
  };

  // 4. Retry up to four times
  for (let attempt = 0; attempt < 4; attempt++) {

    const response = UrlFetchApp.fetch(url, {

      method: "post",

      contentType: "application/json",

      headers: {
        "x-goog-api-key": apiKey
      },

      payload: JSON.stringify(payload),

      muteHttpExceptions: true

    });

    const statusCode = response.getResponseCode();

    const result = JSON.parse(
      response.getContentText()
    );

    // 5. Successful response
    if (statusCode === 200) {

      const aiText = result.candidates?.[0]
        ?.content?.parts
        ?.map(part => part.text || "")
        .join("")
        .trim();

      if (!aiText) {
        throw new Error("Gemini returned an empty response.");
      }

      return aiText;
    }

    // 6. Retry temporary errors
    const temporaryErrors = [429, 500, 502, 503, 504];

    if (!temporaryErrors.includes(statusCode)) {
      throw new Error(
        result.error?.message ||
        "Gemini API failed. HTTP " + statusCode
      );
    }

    Logger.log(
      "Gemini temporary error: " + statusCode +
      ". Attempt " + (attempt + 1)
    );

    // Wait 2, 4, then 8 seconds
    if (attempt < 3) {
      Utilities.sleep(
        Math.pow(2, attempt + 1) * 1000
      );
    }
  }

  throw new Error(
    "Gemini is temporarily overloaded. Please try again later."
  );
}


// 5. EXTRACT TEXT FROM GEMINI RESPONSE
function extractAIText(result) {

  // Check direct output text if available
  if (result.output_text) {
    return result.output_text.trim();
  }

  // Otherwise extract text from model output steps
  const steps = result.steps || [];

  const textParts = [];

  steps.forEach(function(step) {

    if (step.type === "model_output") {

      const content = step.content || [];

      content.forEach(function(part) {

        if (part.type === "text" && part.text) {
          textParts.push(part.text);
        }

      });
    }

  });

  return textParts.join("\n").trim();
}


// 6. CREATE OUTPUT COLUMNS AUTOMATICALLY
function getOrCreateColumn(sheet, header) {

  const lastCol = sheet.getLastColumn();

  const headers = sheet
    .getRange(1, 1, 1, lastCol)
    .getValues()[0];

  const existingIndex = headers.indexOf(header);

  if (existingIndex !== -1) {
    return existingIndex + 1;
  }

  const newColumn = lastCol + 1;

  sheet.getRange(1, newColumn)
    .setValue(header);

  return newColumn;
}


// 7. SET UP AUTOMATIC FORM SUBMISSION TRIGGER
function setupTrigger() {

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error("No connected spreadsheet found.");
  }

  // Remove old triggers for this function
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(function(trigger) {

    if (
      trigger.getHandlerFunction() === "onFormSubmit"
    ) {
      ScriptApp.deleteTrigger(trigger);
    }

  });

  // Create new form submission trigger
  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(spreadsheet)
    .onFormSubmit()
    .create();

  Logger.log("Automation trigger created successfully!");
}


// 8. TEST GEMINI CONNECTION WITHOUT SENDING EMAIL
function testGeminiConnection() {

  const sampleLead = {
    name: "Rahul",
    email: "rahul@example.com",
    phone: "9999999999",
    service: "Website Development",
    budget: "₹10,000–₹30,000",
    requirements:
      "I need a professional business website."
  };

  const response = generateAIResponse(sampleLead);

  Logger.log("Gemini AI Response:");
  Logger.log(response);
}