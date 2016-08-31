# Itineraries
Javascript/SharePoint App for updating People Itineraries in a SharePoint list. Originally created for the use of staff at British Columbia Government and Service Employee's Union.

## Requirements
The app uses the following libraries and their stylesheets:
* JQuery
* Bootstrap 3
* Bootstrap 3 Typeahead
* SP.runtime.js in SharePoint 2013
* SP.js in SharePoint 2013
* SP.taxonomy.js in SharePoint 2013

It also interacts with the following SharePoint list/term stores to perform its functions:
* Itineraries - SharePoint list which houses all itinerary data. The list structure is as follows:
```
Staff: People field; Required
Date: Date field; Required
AM: Text field; Required; Limited to 10 characters
PM: Text field; Required; Limited to 10 characters
```
* Department - Managed Metadata Term Store which stores all department names
* Office - Managed Metadata Term Store which stores all office names
* Local People Results - SharePoint people results in the SiteCollection

## Getting Started

* Place index.html, style.css, app.js in one directory within the SharePoint SiteCollection.
* Prepare your list and term stores as per the required specifications.
* Update the list name and term store hash ids to match those in your SharePoint site.
* Insert a Content Editor WebPart into a SharePoint Page with the link to the index.html file.

## Built With

* Atom editor
* Bootstrap 3
* SharePoint 2013

## Authors

* **Ray Juei-Fu Liu** - [WolfDeNoir](https://github.com/wolfdenoir)

## License

This project is licensed under the GNU GENERAL PUBLIC LICENSE - see the [LICENSE](LICENSE) file for details
