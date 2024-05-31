// Function to create a single-color bar graph for onboarding data over the last five weeks and send it via email
function createAndSendSingleColorOnboardingTrendChart() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Onboarding");
  if (!sheet) {
    throw new Error("Onboarding sheet not found.");
  }

  var today = new Date();

  // Function to get the Wednesday of the week
  function getWednesday(date) {
    var day = date.getDay();
    var wednesday = new Date(date);
    wednesday.setDate(date.getDate() - day + (day >= 3 ? 3 : -4));
    return wednesday;
  }

  // Get the Wednesdays for the last five weeks
  var thisWednesday = getWednesday(today);
  var lastWednesday = new Date(thisWednesday);
  lastWednesday.setDate(thisWednesday.getDate() - 7);
  var twoWeeksAgoWednesday = new Date(lastWednesday);
  twoWeeksAgoWednesday.setDate(lastWednesday.getDate() - 7);
  var threeWeeksAgoWednesday = new Date(twoWeeksAgoWednesday);
  threeWeeksAgoWednesday.setDate(twoWeeksAgoWednesday.getDate() - 7);
  var fourWeeksAgoWednesday = new Date(threeWeeksAgoWednesday);
  fourWeeksAgoWednesday.setDate(threeWeeksAgoWednesday.getDate() - 7);
  var fiveWeeksAgoWednesday = new Date(fourWeeksAgoWednesday);
  fiveWeeksAgoWednesday.setDate(fourWeeksAgoWednesday.getDate() - 7);

  var data = sheet.getDataRange().getValues();

  var weeklyCounts = [0, 0, 0, 0, 0]; // To store the counts for the last five weeks

  data.forEach(function(row) {
    var onboardingDate = new Date(row[7]); // Assuming the date is in the eighth column
    if (onboardingDate >= fiveWeeksAgoWednesday && onboardingDate < fourWeeksAgoWednesday) {
      weeklyCounts[0]++;
    } else if (onboardingDate >= fourWeeksAgoWednesday && onboardingDate < threeWeeksAgoWednesday) {
      weeklyCounts[1]++;
    } else if (onboardingDate >= threeWeeksAgoWednesday && onboardingDate < twoWeeksAgoWednesday) {
      weeklyCounts[2]++;
    } else if (onboardingDate >= twoWeeksAgoWednesday && onboardingDate < lastWednesday) {
      weeklyCounts[3]++;
    } else if (onboardingDate >= lastWednesday && onboardingDate < thisWednesday) {
      weeklyCounts[4]++;
    }
  });

  // Prepare data for the chart without annotations
  var chartData = [
    ['Week', 'Onboardings'],
    [formatDate(fiveWeeksAgoWednesday, 'MM/dd/yyyy') + ' - ' + formatDate(fourWeeksAgoWednesday, 'MM/dd/yyyy'), weeklyCounts[0]],
    [formatDate(fourWeeksAgoWednesday, 'MM/dd/yyyy') + ' - ' + formatDate(threeWeeksAgoWednesday, 'MM/dd/yyyy'), weeklyCounts[1]],
    [formatDate(threeWeeksAgoWednesday, 'MM/dd/yyyy') + ' - ' + formatDate(twoWeeksAgoWednesday, 'MM/dd/yyyy'), weeklyCounts[2]],
    [formatDate(twoWeeksAgoWednesday, 'MM/dd/yyyy') + ' - ' + formatDate(lastWednesday, 'MM/dd/yyyy'), weeklyCounts[3]],
    [formatDate(lastWednesday, 'MM/dd/yyyy') + ' - ' + formatDate(thisWednesday, 'MM/dd/yyyy'), weeklyCounts[4]],
  ];

  // Insert data into a new sheet or a specific range in the existing sheet for chart creation
  var chartSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('OnboardingChart') || SpreadsheetApp.getActiveSpreadsheet().insertSheet('OnboardingChart');
  chartSheet.clear(); // Clear previous data
  chartSheet.getRange(1, 1, chartData.length, chartData[0].length).setValues(chartData);

  var chart = chartSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(chartSheet.getRange(1, 1, chartData.length, chartData[0].length))
    .setPosition(5, 5, 0, 0)
    .setOption('title', 'Onboarding Trend Over the Last Five Weeks')
    .setOption('annotations.alwaysOutside', true)
    .setOption('vAxis', { title: 'Number of Onboardings', minValue: 0 }) // Ensure y-axis starts from 0
    .setOption('hAxis', { title: 'Week' })
    .setOption('colors', ['#16be48']) // Set a single color (e.g., Google Blue) for all bars
    .setOption('legend', { position: 'none' })
    .build();
  chartSheet.insertChart(chart);

  // Get the chart as an image
  var chartBlob = chart.getAs('image/png');

  // Prepare email content
  var emailBody = "Hi Team,<br><br>Please find the onboarding trend for the last five weeks below:<br><br><img src='cid:chartImage'><br><br>Best regards,<br>Your Team";

  // Send the email
  MailApp.sendEmail({
    to: "required mail IDs,",
    subject: "Weekly Onboarding Trend",
    htmlBody: emailBody,
    inlineImages: {
      chartImage: chartBlob
    }
  });

  Logger.log('Onboarding trend chart created and emailed successfully.');
}

// Function to format dates
function formatDate(date, format) {
  var day = date.getDate();
  var month = date.getMonth() + 1;
  var year = date.getFullYear();

  // Add leading zeros if necessary
  if (day < 10) {
    day = '0' + day;
  }
  if (month < 10) {
    month = '0' + month;
  }

  // Replace format placeholders with actual values
  format = format.replace('dd', day);
  format = format.replace('MM', month);
  format = format.replace('yyyy', year);

  return format;
}
