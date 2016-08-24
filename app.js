var arrStaff = [];
var arrTypeahead = [];
var mapTypeahead = {};
var context = new SP.ClientContext.get_current();
var web = context.get_web();
var user = web.get_currentUser();

$(document).ready(function() {
  // On page load, display the current week based on today's date.
  $("#weekNo").text(getWeekOf($("#weekNo").data("offset")));

  /**
    Previous button: Move to previous week

    Process:
    1. Disable form controls
    2. Query itin list for next week's data
    3. On success, enable form controls and populate data
  **/
  $("#prev").on("click",
    function(ev) {
      if ($(ev.currentTarget).hasClass("disabled"))
        return;

      var week = $("#weekNo").data("offset");
      week--;
      $("#weekNo").data("offset", week);
      $("#weekNo").text(getWeekOf(week));
      getItinsJSON(week);
      ev.stopPropagation();
    });

  // Next button: Move to next week
  $("#next").on("click",
    function(ev) {
      if ($(ev.currentTarget).hasClass("disabled"))
        return;

      var week = $("#weekNo").data("offset");
      week++;
      $("#weekNo").data("offset", week);
      $("#weekNo").text(getWeekOf(week));
      getItinsJSON(week);
      ev.stopPropagation();
    });

  /**
    Edit button: Toggles edit controls.

    Process:
    1. Check state of Edit button -
      a. Idle: Change to Editing state and repopulate table with input controls
      b. Editing: Change to Idle state and repopulate table with text labels
  **/
  $("#btnEdit").on("click",
    function(ev) {
      if ($(ev.currentTarget).hasClass("disabled"))
        return;

      if (!isReadyForEdit) {
        alert("Hold your horses dear! Data is still being crunched... Try again soon!");
        return;
      }

      if ($(this).hasClass("btn-default")) {
        $(this).removeClass("btn-default");
        $(this).addClass("btn-primary");
        $(this).html('Editing...');
      } else {
        $(this).removeClass("btn-primary");
        $(this).addClass("btn-default");
        $(this).html('Edit');
      }
      refreshItins();
      ev.stopPropagation();
    });

  context.load(user);
  context.executeQueryAsync(function() {
    console.log(user);
    getTypeaheadTerms();
  }, function() {
    alert("Connection Failed. Refresh the page and try again. :(");
  });
});

function escapeURL(string) {
  return String(string).replace(/[&<>\-'\/]/g, function(s) {
    var entityMap = {
      "&": "%26",
      "<": "%3C",
      ">": "%3E",
      "-": "\\-",
      "'": '%60',
      "\\": "%5C",
      "[": "%5B",
      "]": "%5D"
    };

    return entityMap[s];
  });
}

/**
  Reloads the table with appropriate controls based on
  list of staff and Edit button state. Only empties the
  table content if no staff have been loaded.
**/
function refreshItins() {
  $("#tblItin tbody").empty();

  if (arrStaff.length == 0)
    return;

  for (var i = 0; i < arrStaff.length; i++) {
    var td = $("<td/>", {
      "class": "staff",
      html: '<div>' +
        '<div><img src="' + arrStaff[i].Picture + '"/></div>' +
        '<div><a href="mailto:' + arrStaff[i].Email + '">' + arrStaff[i].Name + '</a>' +
        '<span>' + arrStaff[i].JobTitle + '</span>' +
        '<span>' + arrStaff[i].Phone + '</span></div>' +
        '</div>',
      data: {
        id: (arrStaff[i].ID ? arrStaff[i].ID : -1),
        staff: arrStaff[i].Name
      }
    });
    var td2 = $("<td/>", {
      "class": "itin"
    });
    var tr = $("<tr/>").append(td, td2, td2.clone(), td2.clone(), td2.clone(), td2.clone());

    $("#tblItin tbody").append(tr);
  }

  // If Edit Mode is on, create <input> controls;
  // Otherwise create <span> for viewing only.
  if ($("#btnEdit").hasClass("btn-primary")) {
    var am = $("<div/>", {
      "class": "input-group"
    }).append($("<input/>", {
      maxlength: '10',
      placeholder: "AM",
      "class": "form-control am",
      html: ""
    }));
    var pm = $("<div/>", {
      "class": "input-group"
    }).append($("<input/>", {
      maxlength: '10',
      placeholder: "PM",
      "class": "form-control pm",
      html: ""
    }));
  } else {
    var am = $("<span/>", {
      "class": "am",
      html: ""
    });
    var pm = $("<span/>", {
      "class": "pm",
      html: ""
    });
  }

  $(".itin").append(am, pm);

  $("input.am, input.pm").on("focus", function(ev){
    var btnCopyLeft = $("<span/>", {
      "class": "input-group-btn"
    }).append($("<button/>",{
      type: "button",
      "class": "btn btn-default",
      html: '<span class="glyphicon glyphicon-triangle-left" aria-hidden="true"></span>'
    }));
    var btnCopyRight = $("<span/>", {
      "class": "input-group-btn"
    }).append($("<button/>",{
      type: "button",
      "class": "btn btn-default",
      html: '<span class="glyphicon glyphicon-triangle-right" aria-hidden="true"></span>'
    }));
    if($(this).parent().parent().index() == 1) {
      $(this).parent().append(btnCopyRight);
    } else if($(this).parent().parent().is(":last-child")) {
      $(this).parent().prepend(btnCopyLeft);
    } else {
      $(this).parent().append(btnCopyRight);
      $(this).parent().prepend(btnCopyLeft);
    }
  });

  $("input.am, input.pm").on("blur", function(ev){
    console.log($(this).siblings().length);
    while($(this).siblings().length > 1)
      $(this).siblings().remove();
  });

  // On any change activity, record change in value.
  $("input.am, input.pm").on("change", function(ev) {
    $(this).parent().children().addClass("disabled");
    // Calculate what date it is based on the position of the cell entry in the table.
    var day = getDayOfWeek($("#weekNo").data("offset"), $(this).parent().parent().children().index($(this).parent()));
    var time = $(this).attr("class");
    var staff = $(this).parent().parent().children().eq(0).data("staff");

    var itemId = $(this).parent().data("id");
    var jsonData = {
      "__metadata": {
        "type": "SP.Data.ItinerariesListItem"
      },
      "StaffId": $(this).parent().siblings(":first").data("id"),
      "Date": day.toJSON()
    }

    if ($(this).hasClass('am'))
      jsonData.AM = this.value;
    else
      jsonData.PM = this.value;

    if (itemId != undefined && itemId != '') {
      var strSibling = $(this).siblings().eq(0).val();
      if ((strSibling == null || strSibling == '') &&
        (this.value == null || this.value == '')) {
        // Delete Record
        $.ajax({
          url: _spPageContextInfo.webAbsoluteUrl +
            "/_api/web/lists/GetByTitle('Itineraries')/items(" + itemId + ")",
          method: "POST",
          headers: {
            "X-RequestDigest": $("#__REQUESTDIGEST").val(),
            "accept": "application/json;odata=verbose",
            "content-type": "application/json;odata=verbose",
            "X-HTTP-Method": "DELETE",
            "IF-MATCH": "*"
          },
          context: $(this),
          success: onDeleteJSON,
          error: onQueryFailedJSON
        });
      } else {
        // Update Record
        $.ajax({
          url: _spPageContextInfo.webAbsoluteUrl +
            "/_api/web/lists/GetByTitle('Itineraries')/items(" + itemId + ")",
          method: "POST",
          data: JSON.stringify(jsonData),
          headers: {
            "X-RequestDigest": $("#__REQUESTDIGEST").val(),
            "accept": "application/json;odata=verbose",
            "content-type": "application/json;odata=verbose",
            "X-HTTP-Method": "MERGE",
            "IF-MATCH": "*"
          },
          context: $(this),
          success: onUpdateJSON,
          error: onQueryFailedJSON
        });
      }
    } else if (this.value != null && this.value != '') {
      // Create Record
      $.ajax({
        url: _spPageContextInfo.webAbsoluteUrl +
          "/_api/web/lists/GetByTitle('Itineraries')/items",
        method: "POST",
        data: JSON.stringify(jsonData),
        headers: {
          "X-RequestDigest": $("#__REQUESTDIGEST").val(),
          "accept": "application/json;odata=verbose",
          "content-type": "application/json;odata=verbose"
        },
        context: $(this),
        success: onCreateJSON,
        error: onQueryFailedJSON
      });
    } else { // Input is empty and there is no existing itemId
      console.log("Exception: Input is empty and there is no existing itemId.");
    }

    ev.stopPropagation();
  });

  getItinsJSON($("#weekNo").data("offset"));
}

function onCreateJSON(data) {
  $(this).parent().data("id", data.d.ID);
  $(this).parent().children().removeClass("disabled");
  console.log("itemID " + data.d.ID + " created");
}

function onUpdateJSON(data) {
  console.log("ItemID " + $(this).parent().data('id') + " updated.");
  $(this).parent().children().removeClass("disabled");
}

function onDeleteJSON(data) {
  console.log("ItemID " + $(this).parent().data('id') + " deleted.");
  $(this).parent().removeData();
  $(this).parent().children().removeClass("disabled");
}

// Cleans data from the table without removing the controls
function resetFields() {
  $("#itin-main button, #itin-main input").addClass("disabled");

  $(".itin").removeData();
  if ($("#btnEdit").hasClass("btn-primary"))
    $(".am, .pm").val('');
  else
    $(".am, .pm").html('');
}

function enableFields() {
  $("#itin-main button, #itin-main input").removeClass("disabled");
  $("#itin-splash button").removeClass("disabled");
}

// Requests itin data for the current list of staff in arrStaff.
function getItinsJSON(offset) {
  if (arrStaff.length == 0) {
    return;
  } else {
    resetFields();
  }

  var clauseStaff = "$filter=(Staff eq '" + escapeURL(arrStaff[0].Name) + "'";
  for (var i = 1; i < arrStaff.length; i++) {
    clauseStaff += " or Staff eq '" + escapeURL(arrStaff[i].Name) + "'";
  }

  var searchUrl = _spPageContextInfo.webAbsoluteUrl +
    "/_api/web/lists/GetByTitle('Itineraries')/items?" +
    "$select=ID,StaffId,Staff/Title,Date,AM,PM&$expand=Staff&" +
    clauseStaff + ") and Date ge DateTime'" + getDayOfWeek(offset, 1).toJSON().slice(0, 11) +
    "00:00:00.000Z' and Date le DateTime'" + getDayOfWeek(offset, 5).toJSON().slice(0, 11) +
    "23:59:59.999Z'&$orderby=Staff desc";

  //console.log(searchUrl);
  $.ajax({
    url: searchUrl,
    type: "GET",
    headers: {
      "Accept": "application/json; odata=verbose"
    },
    success: onItinsJSON,
    error: onQueryFailedJSON
  });
}

// Populate the Itin table with the queried data
function onItinsJSON(data) {
  //console.log(data);
  var results = data.d.results;
  if (results.length > 0) {
    $.each(results, function(index, result) {
      var date = new Date(result.Date);
      var index = date.getDay() - 1;
      $(".staff:contains('" + result.Staff.Title + "') ~ .itin:eq(" + index + ")").data("id", result.ID);
      // Verify Edit Mode ON/OFF state and populate the correct controls.
      if ($("#btnEdit").hasClass("btn-primary")) {
        $(".staff:contains('" + result.Staff.Title + "') ~ .itin .am:eq(" + index + ")").val(result.AM);
        $(".staff:contains('" + result.Staff.Title + "') ~ .itin .pm:eq(" + index + ")").val(result.PM);
      } else {
        $(".staff:contains('" + result.Staff.Title + "') ~ .itin .am:eq(" + index + ")").html(result.AM);
        $(".staff:contains('" + result.Staff.Title + "') ~ .itin .pm:eq(" + index + ")").html(result.PM);
      }
    });
  } else {
    console.log("No itin record found.");
  }

  enableFields();
}

function onQueryFailed(sender, args) {
  console.log('Request failed.' + args.get_message() +
    '\n' + args.get_stackTrace());
}

function onQueryFailedJSON(data, errorCode, errorMessage) {
  if (data.responseJSON != null)
    console.log(data.responseJSON.error.code + ': ' + data.responseJSON.error.message.value);
  else
    console.log(data);
}

// Requests staff list of a department or office, or just a single staff(strName)
// from Local People Results
// args:
// strName - name of the department/office/staff
// strType - whether strName is a department or office or staff
function getStaffList(strName, strType) {
  var searchTerm = '';
  var searchUrl = '';
  resetFields();
  if (strType == "staff") {
    searchUrl = _spPageContextInfo.webAbsoluteUrl +
      "/_api/search/query?querytext='PreferredName=\"" + escapeURL(strName) +
      "\"'&selectproperties='PreferredName,PictureURL,JobTitle,WorkPhone,WorkEmail,AccountName'&" +
      "sourceid='B09A7990-05EA-4AF9-81EF-EDFAB16C4E31'";
  } else {
    searchTerm = strType + '="' + strName + '" JobTitle<>"Support" AND (' +
      'JobTitle:"a*" JobTitle:"b*" JobTitle:"c*" JobTitle:"d*" JobTitle:"e*"' +
      'JobTitle:"f*" JobTitle:"g*" JobTitle:"h*" JobTitle:"i*" JobTitle:"j*"' +
      'JobTitle:"k*" JobTitle:"l*" JobTitle:"m*" JobTitle:"n*" JobTitle:"o*"' +
      'JobTitle:"p*" JobTitle:"q*" JobTitle:"r*" JobTitle:"s*" JobTitle:"t*"' +
      'JobTitle:"u*" JobTitle:"v*" JobTitle:"w*" JobTitle:"x*" JobTitle:"y*"' +
      'JobTitle:"z*" JobTitle:"0*" JobTitle:"1*" JobTitle:"2*" JobTitle:"3*"' +
      'JobTitle:"4*" JobTitle:"5*" JobTitle:"6*" JobTitle:"7*" JobTitle:"8*"' +
      'JobTitle:"9*")';
    searchUrl = _spPageContextInfo.webAbsoluteUrl +
      "/_api/search/query?querytext='" + searchTerm +
      "'&sortlist='PreferredName:ascending'&" +
      "selectproperties='PreferredName,PictureURL,JobTitle,WorkPhone,WorkEmail,AccountName'&" +
      "sourceid='B09A7990-05EA-4AF9-81EF-EDFAB16C4E31'&" +
      "rowlimit='100'";
  }
  //console.log(searchUrl);

  $.ajax({
    url: searchUrl,
    type: "GET",
    headers: {
      "Accept": "application/json; odata=verbose"
    },
    success: onStaffListJSON,
    error: onQueryFailedJSON
  });
}

// Writes Staff names into arrStaff and reloads the itin table
function onStaffListJSON(data) {
  arrStaff = [];

  var results = data.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results;
  if (results.length > 0) {
    numReadyForEdit = 0; // Counter for all async calls which has completed retrieving Staff id
    isReadyForEdit = false; // Flag for the Edit button click event to decide whether to swap in input controls.
    $.each(results, function(index, result) {
      var item = {
        'ID': null,
        'Name': result.Cells.results[2].Value,
        'Picture': (result.Cells.results[3].Value != null ? result.Cells.results[3].Value : '/_layouts/15/images/person.gif'),
        'Email': result.Cells.results[6].Value,
        'Phone': (result.Cells.results[5].Value != null ? result.Cells.results[5].Value : ''),
        'JobTitle': result.Cells.results[4].Value,
        'AccountName': result.Cells.results[7].Value
      };

      arrStaff.push(item);

      // This async call retrieves the User ID at the SiteCollection level
      // It also ensures the user is available for record creation
      $.ajax({
        url: _spPageContextInfo.webAbsoluteUrl + "/_api/web/ensureuser",
        type: "POST",
        data: JSON.stringify({
          "logonName": item.AccountName
        }),
        headers: {
          "X-RequestDigest": $("#__REQUESTDIGEST").val(),
          "Accept": "application/json; odata=verbose",
          "content-type": "application/json;odata=verbose"
        },
        indexValue: index,
        success: function(data) {
          arrStaff[index].ID = data.d.Id;
          $(".staff").eq(index).data("id", data.d.Id);
          //console.log(arrStaff[index].ID + ": " + arrStaff[index].Name);
          numReadyForEdit++;
          if (arrStaff.length == numReadyForEdit)
            isReadyForEdit = true;
        },
        error: onQueryFailedJSON
      });
    });

    // Since there are results and the form is currently on splash div,
    // Move the department and office filter controls to the main div.
    // Disable the splash div and enable main div and call the function
    // to reload the itin table.
    if ($("#itin-main").hasClass("hidden")) {
      $("#itin-main").removeClass("hidden");
      $("#itin-splash").addClass("hidden");
      var acStaff = $("#acStaff");
      acStaff.detach();
      acStaff.appendTo($("#container-grpOption-main"));
    }

    refreshItins();

    // No results were found. If the main div is displayed, move the filters
    // to the splash div. Enable splash div and disable main div and display
    // the nothing found message.
  } else if ($("#itin-splash").hasClass("hidden")) {
    $("#itin-splash").removeClass("hidden");
    $("#itin-main").addClass("hidden");
    var acStaff = $("#acStaff");
    acStaff.detach();
    acStaff.appendTo($("#container-grpOption-splash"));
    $("#itin-splash h3").html("Sorry, can't find anything... Try something else.");
    // Since no results were found and the splash div is displayed, just need
    // to display the nothing found message.
  } else {
    $("#itin-splash h3").html("Sorry, can't find anything... Try something else.");
  }

  enableFields();
}

// Returns the week dates in a string.
function getWeekOf(weekNo) {
  var monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  var monday = getDayOfWeek(weekNo, 1);
  var friday = getDayOfWeek(weekNo, 5);

  var weekof = monthNames[monday.getMonth()] + ' ' + monday.getDate() + ' to ' + monthNames[friday.getMonth()] + ' ' + friday.getDate();

  return weekof;
}

// Returns the day of week
function getDayOfWeek(weekNo, dayNo) {
  if (dayNo > 5 || dayNo < 1) {
    jQuery.error("Day number must be between 1 and 5.");
    return;
  }

  var d = new Date();
  d.setDate(d.getDate() + weekNo * 7 - d.getDay() + dayNo);

  return d;
}

function getTypeaheadTerms() {
  getDeptTerms();
  getStaffTerms();

  $('#acStaff .typeahead').typeahead({
    source: arrTypeahead,
    updater: function(item) {
      getStaffList(mapTypeahead[item].name, mapTypeahead[item].type);
      return item;
    }
  });
}

// Requests department list from the Managed Metadata Store
function getDeptTerms() {
  this.session = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
  this.termStore = session.getDefaultSiteCollectionTermStore();
  this.termSet = this.termStore.getTermSet("8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f");
  this.terms = termSet.get_terms();
  context.load(session);
  context.load(termStore);
  context.load(termSet);
  context.load(terms);
  context.executeQueryAsync(Function.createDelegate(this, function(sender, args) {
      var termsEnum = terms.getEnumerator();
      while (termsEnum.moveNext()) {
        var currentTerm = termsEnum.get_current();
        var termName = currentTerm.get_name();
        var item = {
          "name": termName,
          "type": "department"
        }
        mapTypeahead[termName] = item;
        arrTypeahead.push(termName);
      }
      console.log("Department terms loaded.");

      getOfficeTerms();
    }),
    Function.createDelegate(this, this.failedListTaxonomySession));
}

/**
  Requests department list from the Managed Metadata Store.
  Can only be called after the ajax call returns from getDeptTerms() since both
  functions share the same global variables.
**/
function getOfficeTerms() {
  this.termSet = this.termStore.getTermSet("1e0e7cef-a4ea-45c1-aaca-6817c9211330");
  this.terms = termSet.get_terms();
  context.load(session);
  context.load(termStore);
  context.load(termSet);
  context.load(terms);
  context.executeQueryAsync(Function.createDelegate(this, function(sender, args) {
      var termsEnum = terms.getEnumerator();
      while (termsEnum.moveNext()) {
        var currentTerm = termsEnum.get_current();
        var termName = currentTerm.get_name();
        var item = {
          "name": termName,
          "type": "office"
        }
        mapTypeahead[termName] = item;
        arrTypeahead.push(termName);
      }
      console.log("Office terms loaded.");
    }),
    Function.createDelegate(this, this.failedListTaxonomySession));
}

// Requests staff list from local People Results
function getStaffTerms() {
  var searchTerm = 'JobTitle<>"Support" AND (' +
    'JobTitle:"a*" JobTitle:"b*" JobTitle:"c*" JobTitle:"d*" JobTitle:"e*"' +
    'JobTitle:"f*" JobTitle:"g*" JobTitle:"h*" JobTitle:"i*" JobTitle:"j*"' +
    'JobTitle:"k*" JobTitle:"l*" JobTitle:"m*" JobTitle:"n*" JobTitle:"o*"' +
    'JobTitle:"p*" JobTitle:"q*" JobTitle:"r*" JobTitle:"s*" JobTitle:"t*"' +
    'JobTitle:"u*" JobTitle:"v*" JobTitle:"w*" JobTitle:"x*" JobTitle:"y*"' +
    'JobTitle:"z*" JobTitle:"0*" JobTitle:"1*" JobTitle:"2*" JobTitle:"3*"' +
    'JobTitle:"4*" JobTitle:"5*" JobTitle:"6*" JobTitle:"7*" JobTitle:"8*"' +
    'JobTitle:"9*")';
  var searchUrl = _spPageContextInfo.webAbsoluteUrl +
    "/_api/search/query?querytext='" + searchTerm +
    "'&selectproperties='PreferredName,AccountName'&" +
    "sourceid='B09A7990-05EA-4AF9-81EF-EDFAB16C4E31'&" +
    "rowlimit='300'";
  //console.log(searchUrl);
  $.ajax({
    url: searchUrl,
    type: "GET",
    headers: {
      "Accept": "application/json; odata=verbose"
    },
    success: function(data) {
      var results = data.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results;
      if (results.length > 0) {
        $.each(results, function(index, result) {
          var item = {
            "name": result.Cells.results[2].Value,
            "AccountName": result.Cells.results[3].Value,
            "type": "staff"
          };
          mapTypeahead[result.Cells.results[2].Value] = item;
          arrTypeahead.push(result.Cells.results[2].Value);
        });
        console.log("Staff terms loaded.");
      } else {
        console.log("No staff terms found.");
      }
    },
    error: onQueryFailedJSON
  });
}

function recursiveTerms(currentTerm, nestedLoop) {
  // Loop count for formatting purpose.
  var loop = nestedLoop + 1;
  // Get Term child terms
  var terms = currentTerm.get_terms();
  context.load(terms);
  context.executeQueryAsync(
    function() {
      var termsEnum = terms.getEnumerator();
      while (termsEnum.moveNext()) {
        var newCurrentTerm = termsEnum.get_current();
        var termName = newCurrentTerm.get_name();
        termId = newCurrentTerm.get_id();
        // Tab Out format.
        for (var i = 0; i < loop; i++) {
          termsList += "\t";
        }
        termsList += termName + ": " + termId;
        // Check if term has child terms.
        if (currentTerm.get_termsCount() > 0) {
          // Term has sub terms.
          recursiveTerms(newCurrentTerm, loop);
        }
      }
    },
    function() {
      // failure to load terms
    });
}
