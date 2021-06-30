// Created by Max Ostapenko, bvz2001@gmail.com

/*
 * Updates response item in the Google Form with available session time slots
 */
function updateSlots() {
  const spreadsheetId = '1xEd2daCtBSeItYTe3q1-YILtG9fB-MOTOku3Ggm8uIs'; // Google Sheet ID that hosts data for session board
  const sheetName = 'Europe'; // Sessions sheet name

  const sessionsItemId = 1515269446; // ID of the form field with session slots. May change in case the form fields are added manually.

  const slotNameSeparator = ', room ';

  let form = FormApp.getActiveForm();
  let sessionsItem = form.getItemById(sessionsItemId);

  let sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  let sheetsRange = sheet.getRange("A1").getDataRegion();
  let sheetValues = sheetsRange.getValues();

  let cell = sheet.getRange('A1');

  let columnsInfo = {
    speaker: {
      formItemIndex: 0,
    },
    twitter: {
      formItemIndex: 1,
    },
    title: {
      formItemIndex: 2,
    },
    description: {
      formItemIndex: 3,
    },
    type: {
      formItemIndex: 4,
    },
    level: {
      formItemIndex: 5,
    },
    focus: {
      formItemIndex: 6,
    },
    tags: {
      formItemIndex: 7,
    },
    time: {
      'read-only': true,
    },
    room_sponsor: {
      'read-only': true,
    },
    response_id: {},
    response_url: {},
  };

  Object.keys(columnsInfo).forEach(key => {
    columnsInfo[key].column = getColumnIndex(sheetValues[0], key);
  });

  freeSessionSlots = [];

  let formResponses = form.getResponses();
  if (formResponses.length > 0) {
    var formResponse = formResponses[formResponses.length - 1];
    var itemResponses = formResponse.getItemResponses();

    // console.log(formResponse.getResponseForItem(sessionsItem).getResponse().split(slotNameSeparator))

    let responseParsed = formResponse.getResponseForItem(sessionsItem).getResponse().split(slotNameSeparator);
    var timeSubmitted = responseParsed[0];
    var roomSubmitted = responseParsed[1];

    console.log('session submitted: ', [timeSubmitted, roomSubmitted]);
  }

  for (var row = 1; row < sheetValues.length; row++) { // Iterating through sheet rows with sessions
    if (sheetValues[row][columnsInfo.room_sponsor.column] != 'Main Area') { // Skipping sessions in Main Area
      let slotName = sheetValues[row][columnsInfo.time.column] + slotNameSeparator + sheetValues[row][columnsInfo.room_sponsor.column];

      if (sheetValues[row][columnsInfo.speaker.column] == '') {
        // If speaker name is empty then the slot is free

        if (
          formResponse &&
          timeSubmitted && timeSubmitted == sheetValues[row][columnsInfo.time.column] &&
          roomSubmitted && roomSubmitted == sheetValues[row][columnsInfo.room_sponsor.column]
        ) {
          // Fills in the submitted session data to the slot.
          console.log(row, 'free slot:', slotName, 'session matched: ', [timeSubmitted, roomSubmitted]);

          Object.keys(columnsInfo).forEach(key => {
            if (typeof columnsInfo[key].formItemIndex === 'number') {
              cell.offset(row, columnsInfo[key].column).setValue(
                itemResponses[columnsInfo[key].formItemIndex].getResponse()
              );
            }
          });

          cell.offset(row, columnsInfo.response_id.column).setValue(
            formResponse.getId()
          );

          cell.offset(row, columnsInfo.response_url.column).setValue(
            formResponse.getEditResponseUrl()
          );

        } else {
          // Skips empty slot, submitted session doesn't match.
          console.log(row, 'free slot:', slotName, 'added to form field');

          slotReset(columnsInfo, cell, row)

          // Collecting all free slots left for form field update.
          freeSessionSlots.push(slotName);
        }
      } else {
        // Speaker name is filled means the session is booked
        console.log(row, 'slot booked: ', slotName);

        // Removes outdated response data from the sheet
        if (
          formResponse &&
          formResponse.getId() == sheetValues[row][columnsInfo.response_id.column] &&
          ( timeSubmitted != sheetValues[row][columnsInfo.time.column] ||
            roomSubmitted != sheetValues[row][columnsInfo.room_sponsor.column] )
        ) {
          console.log(row, 'free slot', slotName, ', reset session data and added to form field');

          slotReset(columnsInfo, cell, row)

          freeSessionSlots.push(slotName);
        }
      }
    }
  }

  if (freeSessionSlots.length > 0) {
    // Updates session form field with a new list of free slots.
    sessionsItem.asMultipleChoiceItem().setChoiceValues(freeSessionSlots);
  } else {
    // Stops collecting session submissions.
    form.setAcceptingResponses(false);
  }

  return true;
}


/*
 * Gets sheet column index by title
 * @param {array} headerRow
 * @param {string} title
 */
function getColumnIndex(headerRow, title) {
  let index = headerRow.findIndex(element => element == title);

  if (index > -1) {
    return index;
  } else {
    return undefined;
  }
}

function slotReset(columnsInfo, cell, row){
  Object.keys(columnsInfo).forEach(key => {
    if (!columnsInfo[key]['read-only']) {
      cell.offset(row, columnsInfo[key].column).setValue('');
    }
  });
}

function enableAcceptResponses(){
  let form = FormApp.getActiveForm();
  form.setAcceptingResponses(true);
}
