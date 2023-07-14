// Get the current date
var today = new Date();

// Check if today is after day 3 of the month
if (today.getDate() >= 3) {
    // Get the last date of the previous month
    var previousMonth = new Date(today.getFullYear(), today.getMonth(), 0);
    var lastDate = previousMonth.getFullYear() + '-' + (previousMonth.getMonth() + 1) + '-' + previousMonth.getDate();
    console.log("Last date of previous month:", lastDate);
} else {
    // Get the last date of the month of the previous month
    var previousMonth = new Date(today.getFullYear(), today.getMonth() - 1, 0);
    var lastDate = previousMonth.getFullYear() + '-' + (previousMonth.getMonth() + 1) + '-' + previousMonth.getDate();
    console.log("Last date of the month of previous month:", lastDate);
}