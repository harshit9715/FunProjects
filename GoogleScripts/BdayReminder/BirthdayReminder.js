const googleSheetID = 'GOOGLE_SHEET_ID';

/**
 * Event Listner - Handles the requests from Time based Triggers.
 * Accepts events.
 * Check Gmail for birthday emails sent by organization.
 * Get the list of employees whose birthday is today.
 * Check names/emails against list of names or emails of friends in google sheet.
 * Creates a Google Calendar event if there is a birthday today.
 * 
 * 
 * @listner
 * @since   1.0.0
 * @access public
 * 
 * @see     https://developers.google.com/apps-script/reference/calendar/calendar-app#createEvent(String,Date,Date,Object)
 * @see     https://developers.google.com/apps-script/reference/gmail/gmail-app#search(String)
 * @see     https://developers.google.com/apps-script/reference/spreadsheet
 */
function setReminderforFriends() {
    /** @type {string} Subject from which mails are to be searched. */
    let subject = 'Happy Birthday!'; // you can use subject, label and other filters.
    /** @type {number} the script runs every 24 hours; also, we only need birthday mails sent in last 24 hours. */
    let interval = 24 * 60 * 60;
    /** @type {Date} Date object when fetching today's birthday emails. */
    let date = new Date();
    /** @type {number} */
    let timeFrom = Math.floor(date.valueOf() / 1000) - interval;
    /**
     * Function call - Get names/emails of friends from the GoogleSpreadsheet
     *
     * @param Name @type {string} Column name in the spreadsheet, use emails for 100% accuracy.
     * @returns @type {friends}
     */
    let friends = getFriendList("Name");
    /**
     * Function call - GoogleApp.search to query gmail.
     *
     * @param subject @type {string} email subject.
     * @param timeFrom @type {string} emails after specific date-time.
     * @returns @type {object} // email objects 
     */
    let threads = GmailApp.search('subject:' + subject + ' after:' + timeFrom); // Query to search for emails.

    // ofcourse, we get 1 birthday email per day, but if there are seperate emails sent we check names in each.
    for (var i = 0; i < threads.length; i++) {
        /** My organization sends emails "to" employees whose bday it is and all other employees are "cc".
         * We can simply check the "to" for getting the list of employees whose bithday is today.
         * We can compare it with the names or emails of friends in our excel sheet.
         * 
         */
        bdayList = friends.filter(friend => threads[i].getMessages()[0].getTo().toLowerCase().includes(friend.toLowerCase()));
        if (bdayList.length > 0) {
            /** Function call to create a calendar event.
             * @name createBirthdayEvent 
             * @param bdayList @type {list} list of names of friends whose birthday is today.
             */
            createBirthdayEvent(bdayList);
        }
    }
}

// create 1 event for all friends whose birthday is today. 
function createBirthdayEvent(bdayList) {

    /** @type {Date} Event start time */
    let startTime = new Date();
    /** @type {Date} Event end time */
    let endTime = new Date();

    startTime = startTime.setHours(10, 0, 0, 0); // +10:00 to get 10am
    endTime = endTime.setHours(10, 15, 0, 0); // +10:15 to get 10am

    /**
     * Function call - CalendarApp.createEvent to create calendar event.
     *
     * @param subject @type {string} event subject.
     * @param startTime @type {Date} time to start event.
     * @param endTime @type {Date} time to end event.
     * @param options @type {object}
     *      @param description @type {string} "description of event."
     * @returns @type {object} // email objects 
     */
    CalendarApp.createEvent("Birthdays", new Date(startTime), new Date(endTime), {
        description: `Today is ${bdayList.join(", ")} birthday(s).`
    });
}

/**
 * Function call - Get names/emails of friends from google sheet.
 *
 * @name getFriendList
 * @param cName @type {string} name of column to be fetched.
 * @returns @type {list} // string of names or emails. 
 */
function getFriendList(cName) {
    /**
     * Function call - Get names/emails of friends from google sheet.
     *
     * @name getValuesByDataRange
     * @returns @type {list} list of all the values in the google sheet.
     */
    data = getValuesByDataRange();
    val = -1
    data[0].forEach(function (colName, index) {
        if (colName === cName) {
            val = index;
        }
    });
    if (val !== -1) {
        return data.map(item => item[val]).slice(1);
    } else return [];
}


/**
 * Function call - Get all data from a google sheet by sheet ID.
 *
 * @name getValuesByDataRange
 * @returns @type {list} // string of names or emails. 
 */
function getValuesByDataRange() {
    var sheet = SpreadsheetApp.openById(googleSheetID).getSheets()[0];
    let data = sheet.getDataRange().getValues();
    return data;
}