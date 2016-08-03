var arrStaff = [];
var context = new SP.ClientContext.get_current();
var web = context.get_web();
var user = web.get_currentUser();
var oList = web.get_lists().getByTitle('Itineraries');
var oListItem = null;

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
      getItins(week);
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
      getItins(week);
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
    getDeptTerms();
  }, function() {
    alert("Connection Failed. Refresh the page and try again. :(");
  });
});

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
      maxlength: '10',
      placeholder: "AM",
      "class": "form-control am",
      html: ""
    });
    var pm = $("<input/>", {
      maxlength: '10',
      placeholder: "PM",
      "class": "form-control pm",
      html: ""
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
  }

  $(".itin").append(am, pm);

  // On any change activity, record change in value.
  $("input.am, input.pm").on("change", function(ev) {
    $(this).parent().children().addClass("disabled");
    // Calculate what date it is based on the position of the cell entry in the table.
    var day = getDayOfWeek($("#weekNo").data("offset"), $(this).parent().parent().children().index($(this).parent()));
    var time = $(this).attr("class");
    var staff = $(this).parent().parent().children().eq(0).data("staff");

    var itemId = $(this).parent().data("id");
    if (itemId != undefined && itemId != '') {
      var strSibling = $(this).siblings().eq(0).val();
      if ((strSibling == null || strSibling == '') &&
        (this.value == null || this.value == '')) {
        // Delete Record
        oListItem = oList.getItemById(itemId);
        oListItem.deleteObject();
        context.executeQueryAsync(Function.createDelegate(this, onDelete),
          Function.createDelegate(this, onQueryFailed));
      } else {
        // Update Record
        oListItem = oList.getItemById(itemId);
        if ($(this).hasClass('am'))
          oListItem.set_item('AM', this.value);
        else
          oListItem.set_item('PM', this.value);
        oListItem.update();
        context.executeQueryAsync(Function.createDelegate(this, onUpdate),
          Function.createDelegate(this, onQueryFailed));
      }
    } else if (this.value != null && this.value != '') {
      // Create Record
      var itemCreateInfo = new SP.ListItemCreationInformation();
      oListItem = oList.addItem(itemCreateInfo);
      oListItem.set_item('Date', day);
      oListItem.set_item('Staff', SP.FieldUserValue.fromUser(staff));
      if ($(this).hasClass('am'))
        oListItem.set_item('AM', this.value);
      else
        oListItem.set_item('PM', this.value);

      oListItem.update();
      context.load(oListItem);
      context.executeQueryAsync(Function.createDelegate(this, onCreate),
        Function.createDelegate(this, onQueryFailed));
    } else { // Input is empty and there is no existing itemId
      console.log("Exception: Input is empty and there is no existing itemId.");
    }

    ev.stopPropagation();
  });

  getItins($("#weekNo").data("offset"));
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
function getItins(offset) {
  if (arrStaff.length == 0) {
    return;
  } else {
    resetFields();
  }

  var d = new Date();
  var monday = (offset * 7) - d.getDay();
  var friday = (offset * 7) - d.getDay() + 5;
  //console.log("monday: " + monday + " friday: " + friday);
  var clauseStaff = "<In><FieldRef Name='Staff' /><Values>";
  for (var i = 0; i < arrStaff.length; i++) {
    clauseStaff += "<Value Type='User'>" + arrStaff[i].Name + "</Value>";
  }

  clauseStaff += "</Values></In>";

  var camlQuery = new SP.CamlQuery();
  var camlQueryText = "<View>" +
    "<Query>" +
    "<Where>" +
    "<And>" +
    "<And>" +
    "<Leq>" +
    "<FieldRef Name='Date' /> " +
    "<Value IncludeTimeValue='TRUE' Type='DateTime'><Today OffsetDays= '" + friday + "' /></Value>" +
    "</Leq>" +
    "<Geq>" +
    "<FieldRef Name='Date' />" +
    "<Value IncludeTimeValue='TRUE' Type='DateTime'><Today OffsetDays= '" + monday + "' /></Value>" +
    "</Geq>" +
    "</And>" +
    clauseStaff +
    "</And>" +
    "</Where>" +
    "<OrderBy>" +
    "<FieldRef Name='Staff' Ascending='False' />" +
    "</OrderBy>" +
    "</Query>" +
    "</View>";

  //console.log(camlQueryText);
  camlQuery.set_viewXml(camlQueryText);
  this.clcnItinListItems = oList.getItems(camlQuery);
  context.load(clcnItinListItems);
  context.executeQueryAsync(
    Function.createDelegate(this, this.onItins),
    Function.createDelegate(this, this.onQueryFailed)
  );
}

// Populate the Itin table with the queried data
function onItins(sender, args) {
  var count = clcnItinListItems.get_count();
  if (count > 0) {
    var listItemInfo = "";
    var listItemEnumerator = clcnItinListItems.getEnumerator();
    while (listItemEnumerator.moveNext()) {
      var oListItem = listItemEnumerator.get_current();
      var id = oListItem.get_id();
      var am = oListItem.get_item('AM');
      var pm = oListItem.get_item('PM');
      var date = oListItem.get_item('Date');
      var staff = oListItem.get_item('Staff').get_lookupValue();

      $(".staff:contains('" + staff + "') ~ .itin:eq(" + (date.getDay() - 1) + ")").data("id", id);
      // Verify Edit Mode ON/OFF state and populate the correct controls.
      if ($("#btnEdit").hasClass("btn-primary")) {
        $(".staff:contains('" + staff + "') ~ .itin .am:eq(" + (date.getDay() - 1) + ")").val(am);
        $(".staff:contains('" + staff + "') ~ .itin .pm:eq(" + (date.getDay() - 1) + ")").val(pm);
      } else {
        $(".staff:contains('" + staff + "') ~ .itin .am:eq(" + (date.getDay() - 1) + ")").html(am);
        $(".staff:contains('" + staff + "') ~ .itin .pm:eq(" + (date.getDay() - 1) + ")").html(pm);
      }

      //console.log(staff + " - " + date + " - " + date.getDay() + " - AM:" + am + " PM:" + pm);
    }
  } else {
    console.log("No itin record found.");
  }

  enableFields();
}

function onQueryFailed(sender, args) {
  console.log('Request failed.' + args.get_message() +
    '\n' + args.get_stackTrace());
}

function getAutoComplete() {

}

function onAutoComplete() {

}

function getStaff() {

}

// Requests staff list of a department or office(strName)
// from Local People Results
// args:
// strName - name of the department/office chosen
// strType - whether strName is a department or office
function getStaffList(strName, strType) {
  resetFields();
  var searchTerm = strType + '="' + strName + '" JobTitle<>"Support" AND (' +
  'JobTitle:"a*" JobTitle:"b*" JobTitle:"c*" JobTitle:"d*" JobTitle:"e*"' +
  'JobTitle:"f*" JobTitle:"g*" JobTitle:"h*" JobTitle:"i*" JobTitle:"j*"' +
  'JobTitle:"k*" JobTitle:"l*" JobTitle:"m*" JobTitle:"n*" JobTitle:"o*"' +
  'JobTitle:"p*" JobTitle:"q*" JobTitle:"r*" JobTitle:"s*" JobTitle:"t*"' +
  'JobTitle:"u*" JobTitle:"v*" JobTitle:"w*" JobTitle:"x*" JobTitle:"y*"' +
  'JobTitle:"z*" JobTitle:"0*" JobTitle:"1*" JobTitle:"2*" JobTitle:"3*"' +
  'JobTitle:"4*" JobTitle:"5*" JobTitle:"6*" JobTitle:"7*" JobTitle:"8*"' +
  'JobTitle:"9*")';
  var keywordQuery = new Microsoft.SharePoint.Client.Search.Query.KeywordQuery(context);
  var searchUrl = _spPageContextInfo.webAbsoluteUrl +
                  "/_api/search/query?querytext='" + searchTerm +
                  "'&sortlist='PreferredName:ascending'&" +
                  "selectproperties='PreferredName,JobTitle,WorkPhone,WorkEmail'&" +
                  "sourceid='B09A7990-05EA-4AF9-81EF-EDFAB16C4E31'&" +
                  "rowlimit='100'";
  var executor = new SP.RequestExecutor(_spPageContextInfo.webAbsoluteUrl);
  executor.executeAsync(
    {
      url: searchUrl,
      method: "GET",
      headers: { "Accept": "application/json; odata=verbose" },
      success: function (data) {
        var jsonObject = JSON.parse(data.body);
        var results = jsonObject.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results;
        if (results.length == 0) {
          console.log('No related documents were found');
        } else {
          $.each(results, function (index, result) {
              console.log(result);
          });
        }
      },
      error: function (data, errorCode, errorMessage) {
        console.log(errorCode + ': ' + errorMessage);
      }
    }
  );

  // Sorting (Ascending = 0, Descending = 1)
  keywordQuery.set_enableSorting(true);
  var sortproperties = keywordQuery.get_sortList();
  sortproperties.add("PreferredName", 0);
  var properties = keywordQuery.get_selectProperties();
  properties.add("WorkPhone");

  keywordQuery.set_queryText(searchTerm);
  keywordQuery.set_sourceId('B09A7990-05EA-4AF9-81EF-EDFAB16C4E31');
  keywordQuery.set_rowLimit(100);
  keywordQuery.set_trimDuplicates(false);
  var searchExecutor = new Microsoft.SharePoint.Client.Search.Query.SearchExecutor(context);
  results = searchExecutor.executeQuery(keywordQuery);
  context.executeQueryAsync(Function.createDelegate(this, onStaffList),
    Function.createDelegate(this, onQueryFailed));
}

// Writes Staff names into arrStaff and reloads the itin table
function onStaffList() {
  arrStaff = [];

  if (results.m_value.ResultTables[0].ResultRows.length > 0) {
    $.each(results.m_value.ResultTables[0].ResultRows, function() {
      var item = {
        'Name': this.PreferredName,
        'Picture': (this.PictureURL != null?this.PictureURL:'/_layouts/15/images/person.gif'),
        'Email': this.WorkEmail,
        'Phone': (this.WorkPhone != null?this.WorkPhone:''),
        'JobTitle': this.JobTitle
      };
      arrStaff.push(item);
      //console.log(this);
    });

    // Since there are results and the form is currently on splash div,
    // Move the department and office filter controls to the main div.
    // Disable the splash div and enable main div and call the function
    // to reload the itin table.
    if ($("#itin-main").hasClass("hidden")) {
      $("#itin-main").removeClass("hidden");
      $("#itin-splash").addClass("hidden");
      var grpOption = $("#grpOption");
      grpOption.detach();
      grpOption.appendTo($("#container-grpOption-main"));
    }

    refreshItins();

    // No results were found. If the main div is displayed, move the filters
    // to the splash div. Enable splash div and disable main div and display
    // the nothing found message.
  } else if ($("#itin-splash").hasClass("hidden")) {
    $("#itin-splash").removeClass("hidden");
    $("#itin-main").addClass("hidden");
    var grpOption = $("#grpOption");
    grpOption.detach();
    grpOption.appendTo($("#container-grpOption-splash"));
    $("#itin-splash h3").html("What the... Nobody works here? Try some place else.");
    // Since no results were found and the splash div is displayed, just need
    // to display the nothing found message.
  } else {
    $("#itin-splash h3").html("What the... Nobody works here? Try some place else.");
  }

  enableFields();
}

function onCreate(sender, args) {
  $(this).parent().data('id', oListItem.get_id());
  $(this).parent().children().removeClass("disabled");
  console.log("ItemID " + oListItem.get_id() + " created.");
}

function onUpdate(sender, args) {
  $(this).parent().children().removeClass("disabled");
  console.log("ItemID " + $(this).parent().data('id') + " updated.");
}

function onDelete(sender, args) {
  $(this).parent().children().removeClass("disabled");
  console.log("ItemID " + $(this).parent().data('id') + " deleted.");
  $(this).parent().removeData();
  $(this).siblings().val('');
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

// Requests department list from the Managed Metadata Store
function getDeptTerms() {
  this.session = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
  this.termStore = session.getDefaultSiteCollectionTermStore();
  this.termsetDept = this.termStore.getTermSet("8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f");
  this.terms = termsetDept.get_terms();
  context.load(session);
  context.load(termStore);
  context.load(termsetDept);
  context.load(terms);
  context.executeQueryAsync(Function.createDelegate(this, this.onDeptTerms),
    Function.createDelegate(this, this.failedListTaxonomySession));
}

// Requests department list from the Managed Metadata Store
function getOfficeTerms() {
  this.session = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
  this.termStore = session.getDefaultSiteCollectionTermStore();
  this.termsetOffice = this.termStore.getTermSet("1e0e7cef-a4ea-45c1-aaca-6817c9211330");
  this.terms = termsetOffice.get_terms();
  context.load(session);
  context.load(termStore);
  context.load(termsetOffice);
  context.load(terms);
  context.executeQueryAsync(Function.createDelegate(this, this.onOfficeTerms),
    Function.createDelegate(this, this.failedListTaxonomySession));
}

// Populates the Dropdown list with the queried department list.
function onDeptTerms(sender, args) {
  var termsEnum = terms.getEnumerator();
  while (termsEnum.moveNext()) {
    var currentTerm = termsEnum.get_current();
    var termName = currentTerm.get_name();

    $("#ulDept").append(
      $("<li/>").append(
        $("<a/>", {
          href: "#",
          html: termName
        }).on('click', function() {
          $('#btnDept').html($(this).html() + '<span class="caret"></span>');
          $('#btnOffice').html('Office' + '<span class="caret"></span>');
          if (!$('#btnDept').hasClass("btn-info")) {
            $('#btnDept').addClass("btn-info");
            $('#btnOffice').removeClass("btn-info");
          }

          // Trigger Staff List refresh
          getStaffList($(this).html(), 'department');
        })
      )
    );

    if (currentTerm.get_termsCount() > 0) {
      recursiveTerms(currentTerm, 1);
    }
  }

  $("#btnDept").removeClass("disabled");
  getOfficeTerms();
}

// Populates the Dropdown list with the queried department list.
function onOfficeTerms(sender, args) {
  var termsEnum = terms.getEnumerator();
  while (termsEnum.moveNext()) {
    var currentTerm = termsEnum.get_current();
    var termName = currentTerm.get_name();

    $("#ulOffice").append(
      $("<li/>").append(
        $("<a/>", {
          href: "#",
          html: termName
        }).on('click', function() {
          $('#btnOffice').html($(this).html() + '<span class="caret"></span>');
          $('#btnDept').html('Department' + '<span class="caret"></span>');
          if (!$('#btnOffice').hasClass("btn-info")) {
            $('#btnOffice').addClass("btn-info");
            $('#btnDept').removeClass("btn-info");
          }

          // Trigger Staff List refresh
          getStaffList($(this).html(), 'office');
        })
      )
    );

    if (currentTerm.get_termsCount() > 0) {
      recursiveTerms(currentTerm, 1);
    }
  }

  $("#btnOffice").removeClass("disabled");
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
