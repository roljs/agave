/// <reference path="../App.js" />

(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            loadBuildings();
            $('#set-subject').click(setSubject);
            $('#get-subject').click(getSubject);
            $('#add-to-recipients').click(addToRecipients);
        });
    };

    function setSubject() {
        Office.cast.item.toItemCompose(Office.context.mailbox.item).subject.setAsync("Hello world!");
    }

    function getSubject() {
        Office.cast.item.toItemCompose(Office.context.mailbox.item).subject.getAsync(function (result) {
            app.showNotification('The current subject is', result.value)
        });
    }

    function addToRecipients() {
        Office.cast.item.toItemCompose(Office.context.mailbox.item).to.addAsync([Office.context.mailbox.userProfile.emailAddress]);
    }

})();


var buildings = [];
var buildingsSvcUri = 'https://agave.azurewebsites.net/corpbuildings/api/buildings';

function loadBuildings() {
    // Send an AJAX request
    var xhr = $.getJSON(buildingsSvcUri);
    xhr.done(function (data) {

        buildings = data;
        $('#buildings').accordion({ active: false, collapsible: true });

        // On success, 'data' contains a list of buildings.
        $.each(data, function (key, item) {
            // Add a list item for the building.
            $('#buildings').append('<h3><a href="#"' + item.Id + '>' + item.Name + '</a></h3><div><p>' + item.Area + '<span style="display: inline-block; width: 100px;  text-align:right"><button onclick="insertBuildingDetails(' + item.Id + ')">Insert</button></span></p><p>' + item.Address + '</p><img width="60%" src="' + item.MapImageUrl + '"/></div>');

        });

        $('#buildings').accordion('refresh');

    });
}


function insertBuildingDetails(index) {
    var building = buildings[index];
    var buildingDetails = '<div><h4><a href="' + building.DetailsUrl + '">' + building.Name + '</a> - ' + building.Area + '</h4><p>' + building.Address + '</p><img width="40%" src="' + building.MapImageUrl + '"</><br/><a href="' + building.DirectionsUrl + '">Get Directions</a></div>'

    var item = Office.context.mailbox.item;
    item.body.setSelectedDataAsync(buildingDetails, { coercionType: Office.CoercionType.Html });

    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        item.location.setAsync("Microsoft Campus - " + building.Name);
    }
}
