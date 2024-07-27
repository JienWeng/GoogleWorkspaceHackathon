function sendReminderEmails() {
    var salesPipelineSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sales Pipeline");
    var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  
    // Get the number of days from cell E3
    var daysThreshold = settingsSheet.getRange("E3").getValue();
    var today = new Date();
    var lastInteractionIndex = 12;
  
    // Get the sales pipeline data and sales reps data
    var pipelineData = salesPipelineSheet.getDataRange().getValues();
    var repsData = settingsSheet.getRange("G4:I105").getValues();
    
    // Create a map of email addresses for quick lookup
    var emailMap = {};
    repsData.slice(1).forEach(row => {
      if (row[0] && row[1]) {
        emailMap[row[0]] = row[1];
      }
    });
  
    // Loop through the sales pipeline data to check for reminders
    pipelineData.slice(1).forEach(row => {
      if (row[lastInteractionIndex]) {
        var lastInteractionDate = new Date(row[lastInteractionIndex]);
        var daysDifference = Math.floor((today - lastInteractionDate) / (1000 * 60 * 60 * 24));
  
        if (daysDifference > daysThreshold) {
          var [ , companyName, contactName, companyEmail, , , , , , , salesRepName, , , nextStep, ] = row;
          var salesRepEmail = emailMap[salesRepName];
          console.log(salesRepEmail);
          console.log(salesRepName);
  
          if (salesRepEmail) {
            sendEmail(salesRepEmail, `Follow-up Reminder for ${companyName}`, `
              <p>Dear ${salesRepName},</p>
              <p>This is a friendly reminder to follow up with <strong>${contactName}</strong> from <strong>${companyName}</strong> (${companyEmail}). It has been <strong>${daysDifference} days</strong> since the last interaction.</p>
              <p><strong>Next Step:</strong> ${nextStep}</p>
              <p>Please ensure to complete the necessary actions as soon as possible to keep the sales process on track. If you need any assistance or have any questions, do not hesitate to reach out.</p>
              <p>Thank you for your attention to this matter.</p>
              <p>Best regards,<br>Sales Team</p>
            `);
          }
  
          if (nextStep === "Send contract for signing" && companyEmail) {
            sendEmail(companyEmail, "Reminder to Sign and Return the Contract", `
              <p>Dear ${contactName},</p>
              <p>This is a friendly reminder to sign and return the contract sent to <strong>${companyName}</strong>. It has been <strong>${daysDifference} days</strong> since we last interacted, and we are eagerly awaiting your response.</p>
              <p>If you have any questions or need any further assistance, please do not hesitate to reach out.</p>
              <p>Thank you for your attention to this matter.</p>
              <p>Best regards,<br>Sales Team</p>
            `);
          }
        }
      }
    });
  }
  
  // Helper function to send emails
  function sendEmail(to, subject, body) {
    MailApp.sendEmail({ to, subject, htmlBody: body });
  }
  
  function remindTargetProgress() {
    var salesPipelineSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sales Pipeline");
    var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  
    // Get the target value for each team member from the Settings sheet
    var targetData = settingsSheet.getRange("G5:I11").getValues();
    var targetMap = {};
    targetData.forEach(row => targetMap[row[0]] = { email: row[1], target: row[2] });
  
    var today = new Date();
    var monthsLeft = 11 - today.getMonth();
  
    // Get the sales pipeline data
    var pipelineData = salesPipelineSheet.getDataRange().getValues();
  
    // Calculate total sales for each rep
    var salesMap = {};
    pipelineData.slice(1).forEach(row => {
      var salesRepName = row[10];  // Assuming Team member is in the 11th column
      var stage = row[4];  // Assuming Stage is in the 5th column
      var value = row[5];  // Assuming Value is in the 6th column
  
      if (!salesMap[salesRepName]) {
        salesMap[salesRepName] = 0;
      }
  
      if (stage === 'Won') {
        salesMap[salesRepName] += value;
      }
    });
  
    // Send reminder emails to each rep
    Object.keys(targetMap).forEach(salesRepName => {
      var salesRepEmail = targetMap[salesRepName].email;
      var targetValue = targetMap[salesRepName].target;
      var totalSales = salesMap[salesRepName] || 0;
      var amountToReachTarget = targetValue - totalSales;
      var progress = (totalSales / targetValue) * 100;
  
      if (salesRepEmail && amountToReachTarget > 0) {
        sendEmail(salesRepEmail, "Reminder: Yearly Target Progress", `
          <p>Dear ${salesRepName},</p>
          <p>This is a reminder of your progress towards the sales target for this year. You currently have <strong>RM ${totalSales}</strong> in sales.</p>
          <p>You need <strong>RM ${amountToReachTarget}</strong> more to reach your yearly target of <strong>RM ${targetValue}</strong>.</p>
          <p>Your progress towards the target is <strong>${progress.toFixed(2)}%</strong>.</p>
          <p>There are <strong>${monthsLeft}</strong> months left to hit the target.</p>
          <p>Keep pushing to achieve your goal!</p>
          <p>Best regards,<br>Sales Team</p>
        `);
      }else{
        sendEmail(salesRepEmail, "Congratulations on Hitting Your Target!", `
            <p>Dear ${salesRepName},</p>
            <p>Congratulations! You have achieved your sales target for this month with a total of <strong>${totalSales.toFixed(2)}</strong> in sales.</p>
            <p>Your hard work and dedication have paid off. Keep up the excellent work!</p>
            <p>Best regards,<br>Sales Team</p>
          `);
      }
    });
  }
  
  // Function to create a time-driven trigger to run remindTargetProgress on the 1st of every month
  function createMonthlyTrigger() {
    ScriptApp.newTrigger('remindTargetProgress')
      .timeBased()
      .onMonthDay(1)
      .atHour(9) // Adjust the time as needed
      .create();
  }
  
  // Uncomment the line below to create the trigger when you first set up the script
  // createMonthlyTrigger();
  
  
  // Helper function to send emails
  function sendEmail(to, subject, body) {
    MailApp.sendEmail({ to, subject, htmlBody: body });
  }
  // Function to display data entry form
  function toEnterData() {
    showSidebar('DataEntryForm', 'Data Entry Form');
  }
  
  // Function to display update record form
  function toUpdateRecord() {
    showSidebar('UpdateRecord', 'Update Form');
  }
  
  // Function to display delete record form
  function toDeleteRecord() {
    showSidebar('DeleteRecord', 'Delete Form');
  }
  
  // Helper function to show sidebar
  function showSidebar(filename, title) {
    var html = HtmlService.createHtmlOutputFromFile(filename)
        .setTitle(title)
        .setWidth(300);
    SpreadsheetApp.getUi().showSidebar(html);
  }
  
  // Function to handle the opening of the spreadsheet
  function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Data Management')
        .addItem('Data Entry', 'toEnterData')
        .addItem('Update Record', 'toUpdateRecord')
        .addItem('Delete Record', 'toDeleteRecord')
        .addToUi();
  }
  
  // Function to add form data to the Sales Pipeline sheet and notify sales reps of new deals
  function addFormData(formData) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName('Sales Pipeline');
    var settingsSheet = spreadsheet.getSheetByName('Settings');
    
    // Retrieve settings
    var stage = settingsSheet.getRange('A4').getValue();
    var probability = settingsSheet.getRange('B4').getValue();
    var expectedCloseDays = settingsSheet.getRange('E11').getValue();
    
    // Calculate expected revenue and close date
    var expectedRevenue = formData.value * probability;
    var creationDate = new Date();
    var expectedCloseDate = new Date(creationDate.getTime() + expectedCloseDays * 24 * 60 * 60 * 1000);
    
    // Prepare data array
    var data = [
      getNextOrderId(sheet),
      formData.companyName,
      formData.contactName,
      formData.contactEmail,
      stage,
      formData.value,
      probability,
      expectedRevenue,
      creationDate,
      expectedCloseDate,
      formData.teamMember,
      probability,
      creationDate,
      settingsSheet.getRange('K4').getValue(),
      0
    ];
    
    // Write to the sheet
    sheet.appendRow(data);
  
    // Function to get sales rep email
    function getSalesRepEmail(salesRepName) {
      var repsData = settingsSheet.getRange("G4:I105").getValues();
      var emailMap = {};
      repsData.slice(1).forEach(row => emailMap[row[0]] = row[1]);
      return emailMap[salesRepName];
    }
  
    // Notify sales rep of the new deal
    var [orderId, companyName, contactName, companyEmail, stage, value, , , , , salesRepName, , , nextStep,] = data;
    var salesRepEmail = getSalesRepEmail(salesRepName);
  
    if (salesRepEmail) {
      // Send notification email
      sendEmail(salesRepEmail, `New Deal: ${companyName}`, `
        <p>Dear ${salesRepName},</p>
        <p>A new deal has been added for <strong>${companyName}</strong> (${companyEmail}).</p>
        <p>Stage: ${stage}</p>
        <p>Value: ${value}</p>
        <p>Expected Close Date: ${expectedCloseDate}</p>
        <p>Team Member: ${formData.teamMember}</p>
        <p>Next Step: ${nextStep}</p>
        <p>Please review and take necessary actions to progress the deal through the sales pipeline.</p>
        <p>Thank you for your attention to this matter.</p>
        <p>Best regards,<br>Sales Team</p>
      `);
  
       addToCalendar(salesRepEmail, `Close Deal Reminder: ${companyName}`, expectedCloseDate, `
        New deal added for ${companyName}. Please prepare to close the deal by ${expectedCloseDate}.
      `);
    }
  }
  
  // Function to add event to Google Calendar with description
  function addToCalendar(email, eventTitle, eventDate, description) {
    var calendarId = 'primary'; // Replace with your calendar ID if not primary
    var calendar = CalendarApp.getCalendarById(calendarId);
  
    // Create event with description
    var event = calendar.createEvent(eventTitle, eventDate, eventDate, {
      description: description,
      guests: email
    });
  
    Logger.log('Event ID: ' + event.getId());
  }
  
  // Function to get the next order ID
  function getNextOrderId(sheet) {
    var lastRow = sheet.getLastRow();
    var lastOrderId = lastRow > 1 ? sheet.getRange(lastRow, 1).getValue() : 0;
    return lastOrderId + 1;
  }
  
  // Function to get team members from Settings sheet
  function getTeamMembers() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
    return sheet.getRange('G5:G105').getValues().flat().filter(String);
  }
  
  // Function to get dropdown data for stage and next step
  function getDropdownData() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
    var stages = sheet.getRange('A4:A105').getValues().flat().filter(String);
    var nextSteps = sheet.getRange('K4:K105').getValues().flat().filter(String);
    return { stages, nextSteps };
  }
  
  // Function to get product details based on Order ID
  function getProductDetails(orderId) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sales Pipeline');
    var values = sheet.getDataRange().getValues();
    
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] == orderId) {
        return {
          stage: values[i][4],
          value: values[i][5],
          nextStep: values[i][13]
        };
      }
    }
    
    return null; // Order ID not found
  }
  
  // Function to update product details based on form data
  function updateProduct(formData) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sales Pipeline');
    var values = sheet.getDataRange().getValues();
    
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] == formData.orderId) {
        var today = new Date();
        
        // Update fields
        values[i][4] = formData.stage;
        values[i][5] = formData.value;
        values[i][13] = formData.nextStep;
        values[i][11] = today;
        
        // Update the row in the sheet
        sheet.getRange(i + 1, 1, 1, values[i].length).setValues([values[i]]);
        break;
      }
    }
  }
  
  // Function to delete a record based on Order ID
  function deleteRecord(orderId) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sales Pipeline');
    var values = sheet.getDataRange().getValues();
    
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] == orderId) {
        sheet.deleteRow(i + 1);
        return true; // Record deleted successfully
      }
    }
    
    return false; // Order ID not found
  }
  