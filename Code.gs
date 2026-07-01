function doPost(e) {
  try {
    if (!e || !e.postData) {
      return ContentService.createTextOutput(JSON.stringify({"status": "Error: Invalid request"}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var formObject = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var concertSheet   = ss.getSheetByName('Concerts');
    var artistSheet    = ss.getSheetByName('Artists');
    var venueSheet      = ss.getSheetByName('Venues');
    var promoterSheet   = ss.getSheetByName('Promoters');

    // ==========================================
    // 1. CREATE NEW ENTITIES FIRST
    // ==========================================
    // Safely fallback to empty strings so 'undefined' never hits the sheet
    var finalArtistId   = formObject.artistId || "";
    var finalVenueId    = formObject.venueId || "";
    var finalPromoterId = formObject.promoterId || "";

    if ((formObject.isNewArtist === true || formObject.isNewArtist === "true") && formObject.newArtistName) {
      finalArtistId = "A-" + new Date().getTime();
      // Country: a typed value in newArtistCountryOther always wins over the dropdown selection.
      var newArtistCountry = formObject.newArtistCountryOther || formObject.newArtistCountry || "";
      artistSheet.appendRow([finalArtistId, formObject.newArtistName, newArtistCountry, "General"]);
    }

    if ((formObject.isNewVenue === true || formObject.isNewVenue === "true") && formObject.newVenueName) {
      finalVenueId = "V-" + new Date().getTime();
      venueSheet.appendRow([finalVenueId, formObject.newVenueName, formObject.newVenueCity || "", 0, "Venue"]);
    }

    if (promoterSheet && (formObject.isNewPromoter === true || formObject.isNewPromoter === "true") && formObject.newPromoterName) {
      finalPromoterId = "P-" + new Date().getTime();
      promoterSheet.appendRow([finalPromoterId, formObject.newPromoterName]);
    }

    // Commit the new entities to the database before we assign them to concerts
    SpreadsheetApp.flush();

    // Normalize the ticket and review flags, used by both update and create paths below
    var hasTicketValue = (formObject.hasTicket === true || formObject.hasTicket === "true") ? "TRUE" : "FALSE";
    var hasReviewValue = (formObject.hasReview === true || formObject.hasReview === "true") ? "TRUE" : "FALSE";

    // ==========================================
    // 2. CONCERT ACTIONS
    // ==========================================
    if (formObject.action === "delete" && formObject.concertId) {
      var cell = concertSheet.getRange("A:A").createTextFinder(String(formObject.concertId)).matchEntireCell(true).findNext();
      if (cell) {
        concertSheet.deleteRow(cell.getRow());
        return ContentService.createTextOutput(JSON.stringify({"status": "Deleted successfully!"})).setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify({"status": "Error: Concert not found for deletion."})).setMimeType(ContentService.MimeType.JSON);
    }

    if (formObject.action === "update" && formObject.concertId) {
      var cell = concertSheet.getRange("A:A").createTextFinder(String(formObject.concertId)).matchEntireCell(true).findNext();
      if (cell) {
        var row = cell.getRow();
        // 10 columns: A-I as before, plus J = hasTicket. Column I now stores
        // promoterId instead of free-text promoter name.
        if (concertSheet.getMaxColumns() < 12) {
          concertSheet.insertColumnsAfter(concertSheet.getMaxColumns(), 12 - concertSheet.getMaxColumns());
        }

        var safeRowData = [
          formObject.concertId || "",
          formObject.date || "",
          finalArtistId,
          finalVenueId,
          formObject.price || "",
          formObject.currency || "",
          formObject.tour || "",
          formObject.notes || "",
          finalPromoterId,
          hasTicketValue,
          formObject.priceTiers || "",
          hasReviewValue
        ];

        concertSheet.getRange(row, 1, 1, 12).setValues([safeRowData]);
        return ContentService.createTextOutput(JSON.stringify({"status": "Updated successfully!"})).setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify({"status": "Error: Concert not found for update."})).setMimeType(ContentService.MimeType.JSON);
    }

    // ==========================================
    // 3. ARTIST, VENUE & PROMOTER DIRECT ACTIONS
    // ==========================================
    if (formObject.action === "updateArtist" && formObject.artistId) {
      var cell = artistSheet.getRange("A:A").createTextFinder(String(formObject.artistId)).matchEntireCell(true).findNext();
      if (cell) {
        // A typed value in artistCountryOther always wins over the dropdown selection.
        var artistCountry = formObject.artistCountryOther || formObject.artistCountry || "";
        artistSheet.getRange(cell.getRow(), 2, 1, 2).setValues([[formObject.artistName || "", artistCountry]]);
        return ContentService.createTextOutput(JSON.stringify({"status": "Artist updated!"})).setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify({"status": "Error: Artist not found for update."})).setMimeType(ContentService.MimeType.JSON);
    }

    if (formObject.action === "deleteArtist" && formObject.artistId) {
      var cell = artistSheet.getRange("A:A").createTextFinder(String(formObject.artistId)).matchEntireCell(true).findNext();
      if (cell) {
        artistSheet.deleteRow(cell.getRow());
        return ContentService.createTextOutput(JSON.stringify({"status": "Artist deleted!"})).setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify({"status": "Error: Artist not found for deletion."})).setMimeType(ContentService.MimeType.JSON);
    }

    if (formObject.action === "updateVenue" && formObject.venueId) {
      var cell = venueSheet.getRange("A:A").createTextFinder(String(formObject.venueId)).matchEntireCell(true).findNext();
      if (cell) {
        venueSheet.getRange(cell.getRow(), 2, 1, 2).setValues([[formObject.venueName || "", formObject.venueCity || ""]]);
        return ContentService.createTextOutput(JSON.stringify({"status": "Venue updated!"})).setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify({"status": "Error: Venue not found for update."})).setMimeType(ContentService.MimeType.JSON);
    }

    if (formObject.action === "deleteVenue" && formObject.venueId) {
      var cell = venueSheet.getRange("A:A").createTextFinder(String(formObject.venueId)).matchEntireCell(true).findNext();
      if (cell) {
        venueSheet.deleteRow(cell.getRow());
        return ContentService.createTextOutput(JSON.stringify({"status": "Venue deleted!"})).setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify({"status": "Error: Venue not found for deletion."})).setMimeType(ContentService.MimeType.JSON);
    }

    if (formObject.action === "updatePromoter" && formObject.promoterId) {
      if (!promoterSheet) {
        return ContentService.createTextOutput(JSON.stringify({"status": "Error: Promoters sheet not found."})).setMimeType(ContentService.MimeType.JSON);
      }
      var cell = promoterSheet.getRange("A:A").createTextFinder(String(formObject.promoterId)).matchEntireCell(true).findNext();
      if (cell) {
        promoterSheet.getRange(cell.getRow(), 2, 1, 1).setValues([[formObject.promoterName || ""]]);
        return ContentService.createTextOutput(JSON.stringify({"status": "Promoter updated!"})).setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify({"status": "Error: Promoter not found for update."})).setMimeType(ContentService.MimeType.JSON);
    }

    if (formObject.action === "deletePromoter" && formObject.promoterId) {
      if (!promoterSheet) {
        return ContentService.createTextOutput(JSON.stringify({"status": "Error: Promoters sheet not found."})).setMimeType(ContentService.MimeType.JSON);
      }
      var cell = promoterSheet.getRange("A:A").createTextFinder(String(formObject.promoterId)).matchEntireCell(true).findNext();
      if (cell) {
        promoterSheet.deleteRow(cell.getRow());
        return ContentService.createTextOutput(JSON.stringify({"status": "Promoter deleted!"})).setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify({"status": "Error: Promoter not found for deletion."})).setMimeType(ContentService.MimeType.JSON);
    }

    // ==========================================
    // 4. NEW CONCERT CREATION
    // ==========================================
    if (formObject.date && (!formObject.action || formObject.action === "save")) {
      var newConcertId = "C-" + new Date().getTime();
      var safeNewRowData = [
        newConcertId,
        formObject.date || "",
        finalArtistId,
        finalVenueId,
        formObject.price || "",
        formObject.currency || "",
        formObject.tour || "",
        formObject.notes || "",
        finalPromoterId,
        hasTicketValue,
        formObject.priceTiers || "",
        hasReviewValue
      ];
      concertSheet.appendRow(safeNewRowData);
      return ContentService.createTextOutput(JSON.stringify({"status": "Saved successfully!"})).setMimeType(ContentService.MimeType.JSON);
    }

    // Fallback for isolated Artist/Venue/Promoter creations without a concert
    return ContentService.createTextOutput(JSON.stringify({"status": "Entity added successfully!"})).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({"status": "Script Error: " + err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 5. INITIAL DATA LOAD (GET REQUEST)
// ==========================================
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var artistsSheet    = ss.getSheetByName('Artists');
  var venuesSheet     = ss.getSheetByName('Venues');
  var concertsSheet   = ss.getSheetByName('Concerts');
  var promotersSheet  = ss.getSheetByName('Promoters');

  // NOTE: these sheets have TWO header rows (e.g. a title row + a column-label
  // row), so real data starts at row 3, not row 2.
  var artistsData   = artistsSheet.getLastRow()  > 2 ? artistsSheet.getRange(3,  1, artistsSheet.getLastRow()  - 2, 3).getValues() : [];
  var venuesData    = venuesSheet.getLastRow()   > 2 ? venuesSheet.getRange(3,   1, venuesSheet.getLastRow()   - 2, 3).getValues() : [];
  var concertsData  = concertsSheet.getLastRow() > 2 ? concertsSheet.getRange(3, 1, concertsSheet.getLastRow() - 2, concertsSheet.getLastColumn()).getValues() : [];

  // Promoters sheet is optional -- if it doesn't exist yet, return empty list.
  var promotersData = [];
  if (promotersSheet) {
    promotersData = promotersSheet.getLastRow() > 2 ? promotersSheet.getRange(3, 1, promotersSheet.getLastRow() - 2, 2).getValues() : [];
  }

  var result = { artists: artistsData, venues: venuesData, concerts: concertsData, promoters: promotersData };

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
