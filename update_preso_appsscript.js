var SLIDES_ID = 'INSERT_SLIDES_ID';

function main() {
  // Text mappings to change in the presentation
  var mappings = {
    'heading1': 'My presentation ',
    'heading2': 'Last updated: ' +
    Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(),
      'yyyy-MM-dd hh:mm')
  };

  // Tables in the presentation to update from Sheets data.
  var tables = {
    'testtable': {
      id: '<...Sheets ID...>',
      sheetName: 'TableData'
    }
  };
  updatePresentation(SLIDES_ID, mappings, tables);
}

// Prefix used in Slides objectIds to indicate that this object has been renamed
// and is the target for content substitution.
var OBJ_PREFIX = '__plemont';

/**
 * Updates a given presentation, performing:
 *     (0) Object renaming to facilitate repeated updating of objects from data.
 *     (1) Updates to any Sheets-linked charts.
 *     (2) Text updates for any text entities, using the mapping.
 *     (3) Table updates for any table entities, using Sheets as a source.
 * @param {string} id The ID of the presentation.
 * @param {!Object.<string>} mappings Dictionary of text entries to substitute.
 * @param {!Object} tables A dictionary of Sheets to update tables from.
 */
function updatePresentation(id, mappings, tables) {
  // DriveApp.createFile(blob);
  var presentation = getPresentation(id);

  // Create requests for one-time changes to object IDs
  var renameObjectRequests = createTextAndTableRenameRequests(presentation);
  Array.prototype.push.apply(renameObjectRequests,
    createSlideRenameRequests(presentation));
  if (renameObjectRequests.length) {
    batchUpdate(presentation, renameObjectRequests);
    presentation = getPresentation(id);
  }

  // Create requests for changing / refreshing the contents of elements
  var requests = createTextReplacementRequests(presentation, mappings, tables);
  Array.prototype.push.apply(requests,
    createRefreshSheetsChartsRequests(presentation));
  batchUpdate(presentation, requests);
}

/**
 * Creates requests to rename any slides in the deck if they do not conform to
 * the naming convention required for the Chrome extension autoplay hack.
 * @param {!Object} presentation The Slides presentation object.
 * @return {!Array.<Request>} rename requests.
 */
function createSlideRenameRequests(presentation) {
  var newIdRequests = [];
  var slides = presentation.slides;
  var totalSlides = slides.length;
  for (var i = 0; i < slides.length; i++) {
    var slide = slides[i];
    var newObjectId = [OBJ_PREFIX, i, totalSlides].join('_');
    if (slide.objectId !== newObjectId) {
      Array.prototype.push.apply(newIdRequests,
        createRenameObjectRequests(slide, newObjectId));
    }
  }
  return newIdRequests;
}

/**
 * Retrieves a presentation.
 * @param {string} presentationId The ID of the presentation to retrieve.
 * @return {!Object} The object representing the presentation.
 */
function getPresentation(presentationId) {
  var options = {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    }
  };
  var url = 'https://slides.googleapis.com/v1/presentations/' + presentationId;
  var response = UrlFetchApp.fetch(url, options);
  return JSON.parse(response);
}

/**
 * Replaces text in text components or tables, where the objectId indicates that
 * a substitution should take place.
 * @param {!Object} presentation The presentation object.
 * @param {!Object.<string>} textMappings Dictionary of text mappings.
 * @param {!Object} tableMappings Dictionary of table mappings.
 * @return {!Array.<Request>}
 */
function createTextReplacementRequests(presentation, textMappings,
    tableMappings) {
  var requests = [];
  var slides = presentation.slides;
  for (var i = 0; i < slides.length; i++) {
    var slide = slides[i];
    var pageElements = slide.pageElements;
    for (var j = 0; j < pageElements.length; j++) {
      var pageElement = pageElements[j];
      // Determine first whether the objectId indicates that this object
      // requires text or tables to be updates
      if (isObjectForTextSub(pageElement)) {
        // The objectId contains the key for either the text or table mapping
        var key = getKeyFromPageElement(pageElement);
        // Determine whether the object is a text object
        if (pageElement.shape && pageElement.shape.text) {
          // Replacing text consists of deleting the old text object and
          // inserting a new.
          if (textMappings[key]) {
            requests.push(createDeleteTextRequest(pageElement));
            requests.push(createInsertTextRequest(pageElement,
              textMappings[key]));
          }
        } else if (pageElement.table) {
          // If instead the object is a table, then replace each cell of the
          // table with data from a spreadsheet, if a mapping exists.
          var spreadsheetInfo = tableMappings[key];
          if (spreadsheetInfo) {
            // It is necessary to obtain a 2d array, showing both the dimensions
            // of the target table, to determine the dimensions to be requested
            // from the Sheet, and also the array shows whether any cell is
            // empty, as attempting to delete all existing text from an empty
            // cell causes an error.
            var tableDimensions = getTableDimensions(pageElement);
            var newTable = loadTableFromSpreadsheet(spreadsheetInfo,
              tableDimensions);
            for (var m = 0; m < newTable.length; m++) {
              var row = newTable[m];
              for (var n = 0; n < row.length; n++) {
                if (tableDimensions[m][n]) {
                  requests.push(
                    createDeleteTableTextRequests(pageElement, m, n));
                }
                requests.push(createInsertTableTextRequests(
                  pageElement, m, n, row[n]));
              }
            }
          }
        }
      }
    }
  }
  return requests;
}

/**
 * Loads data from a specific Sheet in a spreadsheet, of the dimensions
 * specified.
 * @param {!Object} spreadsheetInfo An object containing the ID of the
 *     Spreadsheet and the name of the Sheet.
 * @param {!Array.<!Array.<*>>} tableDimensions A 2D array, of the same
 *     dimensions as required for data to be retrieved.
 * @return {!Array.<!Array.<!Object>>} 2d array of data from spreadsheet.
 */
function loadTableFromSpreadsheet(spreadsheetInfo, tableDimensions) {
  return SpreadsheetApp
      .openById(spreadsheetInfo.id)
      .getSheetByName(spreadsheetInfo.sheetName)
      .getRange(1, 1, tableDimensions.length, tableDimensions[0].length)
      .getValues();
}

/**
 * Builds a 2d array representing the dimensions of a table for a given element
 * on the Slides page. Each array element contains true if the corresponding
 * cell has content, and false if it is empty.
 * @param {!Object} pageElement The table object from the page.
 * @return {!Array.<!Array.<boolean>>}
 */
function getTableDimensions(pageElement) {
  var data = [];
  var rows = pageElement.table.tableRows;
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var newRow = [];
    for (var j = 0; j < row.tableCells.length; j++) {
      var cell = row.tableCells[j];
      newRow.push(cell.text ? true : false);
    }
    data.push(newRow);
  }
  return data;
}

/**
 * Creates a request for use with batchUpdate to delete all the text in a given
 * element on the page.
 * @param {!Object} pageElement The element on the page to delete text from.
 * @return {!Object} The request object.
 */
function createDeleteTextRequest(pageElement) {
  return {
    deleteText: {
      objectId: pageElement.objectId,
      textRange: {
        type: 'ALL'
      },
    }
  };
}

/**
 * Creates a request for use with batchUpdate to insert text into a given
 * element on the page.
 * @param {!Object} pageElement The element on the page to delete text from.
 * @param {string} The text to insert.
 * @return {!Object} The request object.
 */
function createInsertTextRequest(pageElement, text) {
  return {
    insertText: {
      objectId: pageElement.objectId,
      text: text,
      insertionIndex: 0
    }
  };
}

/**
 * Creates a request for use with batchUpdate to delete text from a given cell
 * in a table
 * @param {!Object} pageElement The table on the page to delete text from.
 * @param {number} rowIndex The row index.
 * @param {number} colIndex The column index.
 * @return {!Object} The request object.
 */
function createDeleteTableTextRequests(pageElement, rowIndex, colIndex) {
  return {
    deleteText: {
      objectId: pageElement.objectId,
      cellLocation: {
        rowIndex: rowIndex,
        columnIndex: colIndex
      },
      textRange: {
        type: 'ALL',
      }
    }
  };
}

/**
 * Creates a request for use with batchUpdate to insert text into a given cell
 * in a table
 * @param {!Object} pageElement The table on the page to delete text from.
 * @param {number} rowIndex The row index.
 * @param {number} colIndex The column index.
 * @return {!Object} The request object.
 */
function createInsertTableTextRequests(pageElement, rowIndex, colIndex, text) {
  return {
    insertText: {
      objectId: pageElement.objectId,
      cellLocation: {
        rowIndex: rowIndex,
        columnIndex: colIndex
      },
      text: text,
      insertionIndex: 0
    }
  };
}

/**
 * Extracts the key from a pageElement object ID. For example, object ID:
 * __plemont_1_2_3_headingText returns a key of 'headingText'
 * @param {!Object} pageElement The object from which to extract the key.
 * @return {?string} The key, or null if none was found.
 */
function getKeyFromPageElement(pageElement) {
  // var r = /\_([^_]+)$/;
  var r = new RegExp(OBJ_PREFIX + '_\\d+_\\d+_\\d+_(.*)$');
  var matches = r.exec(pageElement.objectId);
  if (matches && matches.length) {
    return matches[1];
  }
}

/**
 * Determines whether the object on the page expects text replacement.
 * @param {!Object} pageElement The element to test.
 * @return {boolean}
 */
function isObjectForTextSub(pageElement) {
  return pageElement.objectId.substring(0, OBJ_PREFIX.length) === OBJ_PREFIX;
}

/**
 * Retrieves an array of TextElements from the specified object, if it is a
 * table.
 * @param {!Object} pageElement The table element.
 * @return {?Array.<!TextElement>}
 */
function extractTableTopLeftTextElements(pageElement) {
  if (pageElement.table && pageElement.table.tableRows &&
      pageElement.table.tableRows[0].tableCells &&
      pageElement.table.tableRows[0].tableCells[0] &&
      pageElement.table.tableRows[0].tableCells[0].text &&
      pageElement.table.tableRows[0].tableCells[0].text.textElements) {
    return pageElement.table.tableRows[0].tableCells[0].text.textElements;
  }
}

/**
 * Retrieves an array of TextElements from the specified object, if it is a
 * shape containing text.
 * @param {!Object} pageElement The shape element.
 * @return {?Array.<!TextElement>}
 */
function extractLabelTextElements(pageElement) {
  if (pageElement.shape && pageElement.shape.text &&
      pageElement.shape.text.textElements) {
    return pageElement.shape.text.textElements;
  }
}

/**
 * Creates the necessary requests for use with batchUpdate, to rename elements
 * in the presentation, where text substitution markers e.g. ${name} are found.
 * For example, a shape with text "${name}" will have its objectId changes to
 * something like "<prefix>_name".
 * @param {!Object} presentation The presentation object.
 * @return {!Array.<Request>}
 */
function createTextAndTableRenameRequests(presentation) {
  var regex = /^\$\{.*\}\n$/;
  var slides = presentation.slides;
  var requests = [];
  for (var i = 0; i < slides.length; i++) {
    var slide = slides[i];
    var pageElements = slide.pageElements;
    for (var j = 0; j < pageElements.length; j++) {
      var pageElement = pageElements[j];
      // Test to see whether the element is already named with the prefix.
      if (!isObjectForTextSub(pageElement)) {
        // Extract textElements from either table or text.
        var textElements = extractTableTopLeftTextElements(pageElement) ||
          extractLabelTextElements(pageElement);
        if (textElements) {
          for (var k = 0; k < textElements.length; k++) {
            var textElement = textElements[k];
            // Test to see whether the text is of the form ${...}
            if (textElement.textRun && textElement.textRun.content &&
              regex.test(textElement.textRun.content)) {
              var content = textElement.textRun.content;
              // Create the new object and remove the old
              var newObjId = [OBJ_PREFIX, i, j, k,
                content.substring(2, content.length - 2)].join('_');
              var pair = createRenameObjectRequests(pageElement, newObjId);
              Array.prototype.push.apply(requests, pair);
            }
          }
        }
      }
    }
  }
  return requests;
}

/**
 * Creates a pair of request objects needed to effectively rename an object.
 * This is achieved by duplicating the required object with a new name, and
 * deleting the original object
 * @param {!Object} pageElement The object to be renamed.
 * @param {string} newObjId The ID to rename to.
 * @return {!Array.<Request>} A pair of requests.
 */
function createRenameObjectRequests(pageElement, newObjId) {
  var duplicateRequest = {
    duplicateObject: {
      objectId: pageElement.objectId,
      objectIds: {}
    }
  };
  duplicateRequest.duplicateObject.objectIds[pageElement.objectId] = newObjId;
  var deleteRequest = {
    deleteObject: {
      objectId: pageElement.objectId
    }
  };
  return [duplicateRequest, deleteRequest];
}

/**
 * Creates a list of request objects for refreshing any Sheets-linked charts
 * in the presentation.
 * @param {!Object} presentation The presentation.
 * @return {!Array.<Request>}
 */
function createRefreshSheetsChartsRequests(presentation) {
  var objectIds = [];
  var slides = presentation.slides;
  for (var i = 0; i < slides.length; i++) {
    var slide = slides[i];
    var pageElements = slide.pageElements;
    for (var j = 0; j < pageElements.length; j++) {
      var pageElement = pageElements[j];
      if (pageElement.sheetsChart) {
        objectIds.push(pageElement.objectId);
      }
    }
  }
  return objectIds.map(function(objectId) {
    return {refreshSheetsChart: {objectId: objectId}};
  });
}

/**
 * Sends modification requests to Slides API
 * @param {!Object} presentation The presentation to update.
 * @param {!Array.<Request>} The requests to send.
 */
function batchUpdate(presentation, requests) {
  if (!requests.length) {
    return;
  }
  var options = {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    method: 'POST',
    payload: JSON.stringify({
      requests: requests
    }),
    contentType: 'application/json'
  };
  var url = 'https://slides.googleapis.com/v1/presentations/' +
    presentation.presentationId + ':batchUpdate';
  UrlFetchApp.fetch(url, options);
}
