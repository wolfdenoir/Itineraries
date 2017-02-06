var _ITIN_LIST_NAME = 'Itineraries';
var _DEPT_TERMSTORE_ID = '8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f';
var _OFFICE_TERMSTORE_ID = '3588dfe5-e02f-4631-b768-ce6137ebbb61';
var _LOCAL_PEOPLE_RESULT_ID = 'B09A7990-05EA-4AF9-81EF-EDFAB16C4E31';
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

      if (!isLastItem) {
        alert("Hold your horses dear! Data is still being crunched... Try again soon!");
        return;
      }

      if ($(this).hasClass("btn-default")) {
        $(this).removeClass("btn-default");
        $(this).addClass("btn-primary");
        $(this).html('<span class="glyphicon glyphicon-pencil" aria-hidden="true"></span>&nbsp;Editing');
      } else {
        $(this).removeClass("btn-primary");
        $(this).addClass("btn-default");
        $(this).html('<span class="glyphicon glyphicon-eye-open" aria-hidden="true"></span>&nbsp;Viewing');
      }
      refreshItins();
      ev.stopPropagation();
    });

  // If the window has been resized, remove CellCopy buttons and
  // remove static width from the containing div.
  $(window).resize(function() {
    if ($(".cellcopy").length > 0) {
      $(".itin").removeAttr('style');
      $(".cellcopy").remove();
    }
  });

  context.load(user);
  context.executeQueryAsync(function() {
    console.log(user);
    getTypeaheadTerms();
  }, function() {
    alert("Connection Failed. Refresh the page and try again. :(");
  });

  $('[data-toggle="tooltip"]').tooltip();
});

function escapeURL(string) {
  return String(string).replace(/[&<>\-'\/]/g, function(s) {
    var entityMap = {
      "&": "%26",
      "<": "%3C",
      ">": "%3E",
      "-": "-",
      "'": "''",
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
    var am = $("<input/>", {
      maxlength: '20',
      placeholder: "AM",
      "class": "form-control am",
      html: ""
    });
    var pm = $("<input/>", {
      maxlength: '20',
      placeholder: "PM",
      "class": "form-control pm",
      html: ""
    });
    var group = $("<div/>", {
      "class": "input-group"
    }).append(am, pm);
    $(".itin").append(group);
    $(".itin .input-group").each(function(index) {
      $(this).attr("tabindex", $("input, button, a").length + index);
    });
  } else {
    var am = $("<span/>", {
      "class": "am",
      html: ""
    });
    var pm = $("<span/>", {
      "class": "pm",
      html: ""
    });
    $(".itin").append(am, pm);
  }

  var lastIndex;

  $(".itin .input-group").on("focusin", function(ev) {
    // Since the event fires every time an element is focused in the div,
    // We only need to handle the first focusin event.
    if (parseInt($(this).attr("tabindex")) == lastIndex)
      return;
    else
      lastIndex = parseInt($(this).attr("tabindex"));

    // This removes any previously active cellcopy buttons and static width
    while ($(".cellcopy").length > 0) {
      $(".itin").removeAttr('style');
      $(".cellcopy").remove();
    }

    var btnCopyLeft = $("<span/>", {
      "class": "input-group-btn cellcopy"
    }).append($("<button/>", {
      type: "button",
      "class": "btn btn-default",
      "data-toggle": "tooltip",
      "data-placement": "top",
      title: "Paste Left",
      tabindex: $("input, button, a").length,
      html: '<span class="glyphicon glyphicon-triangle-left" aria-hidden="true"></span>'
    }));

    btnCopyLeft.on("click", function(ev) {
      copyToCells(this, false);
    });

    var btnCopyRight = $("<span/>", {
      "class": "input-group-btn cellcopy"
    }).append($("<button/>", {
      type: "button",
      "class": "btn btn-default",
      "data-toggle": "tooltip",
      "data-placement": "top",
      title: "Paste Right",
      tabindex: $("input, button, a").length + 1,
      html: '<span class="glyphicon glyphicon-triangle-right" aria-hidden="true"></span>'
    }));

    btnCopyRight.on("click", function(ev) {
      copyToCells(this);
    });

    $(this).parent().width($(this).parent().width());

    if ($(this).parent().index() == 1) {
      $(this).append(btnCopyRight);
    } else if ($(this).parent().is(":last-child")) {
      $(this).prepend(btnCopyLeft);
    } else {
      $(this).append(btnCopyRight);
      $(this).prepend(btnCopyLeft);
    }

    $('[data-toggle="tooltip"]').tooltip();
    ev.stopPropagation();
  });

  // On any change activity, record change in value.
  $("input.am, input.pm").on("change", editItin);

  getItinsJSON($("#weekNo").data("offset"));
}

function editItin(ev) {
  $(this).parent().children().addClass("disabled");
  // Calculate what date it is based on the position of the containing <td> in the table.
  var day = getDayOfWeek($("#weekNo").data("offset"), $(this).parent().parent().index());
  var time = $(this).attr("class");
  var staff = $(this).parent().parent().siblings(":first").data("staff");

  var itemId = $(this).parent().parent().data("id");
  var jsonData = {
    "__metadata": {
      "type": "SP.Data.ItinerariesListItem"
    },
    "StaffId": $(this).parent().parent().siblings(":first").data("id"),
    "Date": day.toJSON()
  }

  var strSibling = $(this).siblings("input").val();
  if ($(this).hasClass('am')) {
    jsonData.AM = this.value;
    jsonData.PM = strSibling;
  } else {
    jsonData.PM = this.value;
    jsonData.AM = strSibling;
  }

  var listURL = _spPageContextInfo.webAbsoluteUrl +
    "/_api/web/lists/GetByTitle('" + _ITIN_LIST_NAME + "')/items";
  if (itemId != undefined && itemId != '') {
    if ((strSibling == null || strSibling == '') &&
      (this.value == null || this.value == '')) {
      // Delete Record
      $.ajax({
        url: listURL + "(" + itemId + ")",
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
        url: listURL + "(" + itemId + ")",
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
      url: listURL,
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
    // do nothing
    console.log("Warning: Input is empty and there is no existing itemId.");
  }

  ev.stopPropagation();
}

/**
  Save the current input value across other inputs in the left or right direction.
**/
function copyToCells(target, isDirectedRight) {
  var isDirectedRight = (isDirectedRight === null || isDirectedRight === undefined) ? true : isDirectedRight;
  if ($(target).parent().parent().index() == 1 && !isDirectedRight ||
    $(target).parent().parent().is(":last-child") && isDirectedRight) {
    console.log("Exeption: There are no inputs in the specified direction.");
    return;
  }

  var am = $(target).siblings(".am").val();
  var pm = $(target).siblings(".pm").val();

  if (isDirectedRight) {
    $(target).parent().parent().nextAll().children().each(function(index) {
      $(this).children(".am").val(am);
      $(this).children(".pm").val(pm);
      $(this).children(".am").trigger("change");
    });
  } else {
    $(target).parent().parent().prevUntil(".staff").children().each(function(index) {
      $(this).children(".am").val(am);
      $(this).children(".pm").val(pm);
      $(this).children(".am").trigger("change");
    });
  }
}

function onCreateJSON(data) {
  $(this).parent().parent().data("id", data.d.ID);
  $(this).parent().children().removeClass("disabled");
  console.log("itemID " + data.d.ID + " created");
}

function onUpdateJSON(data) {
  console.log("ItemID " + $(this).parent().parent().data('id') + " updated.");
  $(this).parent().children().removeClass("disabled");
}

function onDeleteJSON(data) {
  console.log("ItemID " + $(this).parent().parent().data('id') + " deleted.");
  $(this).parent().parent().removeData();
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
    "/_api/web/lists/GetByTitle('" + _ITIN_LIST_NAME + "')/items?" +
    "$select=ID,StaffId,Staff/Title,Date,AM,PM&$expand=Staff&" +
    clauseStaff + ") and Date ge DateTime'" + getDayOfWeek(offset, 1).toJSON().slice(0, 11) +
    "00:00:00.000Z' and Date le DateTime'" + getDayOfWeek(offset, 5).toJSON().slice(0, 11) +
    "23:59:59.999Z'&$orderby=Staff desc&$top=200";

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
  var results = data.d.results;
  if (results.length > 0) {
    //console.log(results);
    $.each(results, function(index, result) {
      var date = new Date(result.Date);
      var index = date.getDay() - 1;
      var staff = String(result.Staff.Title).replace(/'/g, "\\'");

      $(".staff:contains('" + staff + "') ~ .itin:eq(" + index + ")").data("id", result.ID);
      // Verify Edit Mode ON/OFF state and populate the correct controls.
      if ($("#btnEdit").hasClass("btn-primary")) {
        $(".staff:contains('" + staff + "') ~ .itin .am:eq(" + index + ")").val(result.AM);
        $(".staff:contains('" + staff + "') ~ .itin .pm:eq(" + index + ")").val(result.PM);
      } else {
        $(".staff:contains('" + staff + "') ~ .itin .am:eq(" + index + ")").html(result.AM);
        $(".staff:contains('" + staff + "') ~ .itin .pm:eq(" + index + ")").html(result.PM);
      }
    });
  } else {
    console.log("No itin record found.");
  }

  enableFields();
}

function onQueryFailedJSON(data, errorCode, errorMessage) {
  if (data.responseJSON != null)
    console.log(data.responseJSON.error.code + ': ' + data.responseJSON.error.message.value);
  else
    console.log(data);

  alert("There was a problem updating the data. Please refresh the page and try agin.");
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
      "/_api/search/query?querytext='PreferredName=\"" + String(escapeURL(strName)).replace(/-/g, "\\-") +
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
    isLastItem = false; // Flag for checking if the last item is being processed.
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
        node: item,
        index: index,
        success: function(data) {
          item.ID = data.d.Id;
          arrStaff.push(item);
          //console.log(item.ID + ": " + item.Name);
          if (results.length == (index + 1)) {
            isLastItem = true;
            refreshItins();
          }
        },
        error: function(data, errorCode, errorMessage) {
          if (data.responseJSON != null)
            console.log(data.responseJSON.error.code + ': ' + data.responseJSON.error.message.value);
          else
            console.log(data);

          if (results.length == (index + 1)) {
            isLastItem = true;
            refreshItins();
          }
        }
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
  this.termSet = this.termStore.getTermSet(_DEPT_TERMSTORE_ID);
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
  this.termSet = this.termStore.getTermSet(_OFFICE_TERMSTORE_ID);
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
    "sourceid='" + _LOCAL_PEOPLE_RESULT_ID + "'&" +
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
