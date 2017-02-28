# Itineraries
Javascript/SharePoint web app for updating People Itineraries in a SharePoint list. Originally created for the use of staff at British Columbia Government and Service Employee's Union(BCGEU). The Itineraries is an existing app BCGEU uses to update its staff's activities and view them per week's glance. This is a redevelopment/update of the app ever since the company has employed SharePoint for its intranet.

Contrary to a SharePoint app built from Visual Studio, this app is built solely from HTML and Javascript. I built it this way so I can quickly deploy changes to the app and bring it live, versus having to go through solution deployment each time I make changes. It does mean there are manual work involved in setting up the data backend and configuration, but I consider these minimal and it gives you a good way to learn how the app works in the background.

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
AM: Text field; Required; Limited to 20 characters
PM: Text field; Required; Limited to 20 characters
```
* STAT Holidays - SharePoint list which stores information on the authorized day off for the company. The list structure is as follows:
```
Title: The name of the STAT Holiday
Date: The date of the STAT Holiday
```
* Department - Managed Metadata Term Store which stores all department names
* Office - Managed Metadata Term Store which stores all office names
* Local People Results - SharePoint people results in the SiteCollection

## Getting Started

* Place index.html, style.min.css, app.min.js in one directory within the SharePoint SiteCollection. Update the script paths in index.html accordingly
* Prepare your lists and term stores as per the specifications
* Update list IDs and term store IDs to match those in your SharePoint site
* Create a SharePoint page where you wish to deploy the app
* On this page, insert a Content Editor WebPart with link to the index.html file

## Built With

* Atom editor
* Bootstrap 3
* SharePoint 2013

## Authors

* **Ray Juei-Fu Liu** - [WolfDeNoir](https://github.com/wolfdenoir)

## License

This project is licensed under the GNU GENERAL PUBLIC LICENSE - see the [LICENSE](LICENSE) file for details
