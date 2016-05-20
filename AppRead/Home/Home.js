/// <reference path="../App.js" />
// global app

(function () {
    'use strict';

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            Office.context.mailbox.item.body.getAsync(
			  "text",
			  { asyncContext:"This is passed to the callback" },
			  function callback(result) {
			    // Do something with the result
				var str = result.value; 
				var res = str.match(/Workflow\sId:\s[0-9|a-z|A-Z]{6}/g).toString();      				
				var wf = res.replace(/Workflow Id:\s/g, "WorkflowId=");
				
				$("#content-main").html("The RegEx matched is---> "+res)
				
				//You can then change it up to add to an IFrame, or an angular application route
				//var url = "https://xomino365.azurewebsites.net/yourapplication?&"+wf + "&"
				//$("#content-main").append("<iframe style='width: 100%; overflow: hidden' src='"+url+"' frameborder=0/>")
			  });
        });
    };

    // Displays the "Subject" and "From" fields, based on the current mail item
    function displayItemDetails() {
        var item = Office.cast.item.toItemRead(Office.context.mailbox.item);
        $('#subject').text(item.subject);

        var from;
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            from = Office.cast.item.toMessageRead(item).from;
        } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            from = Office.cast.item.toAppointmentRead(item).organizer;
        }

        if (from) {
            $('#from').text(from.displayName);
            $('#from').click(function () {
                app.showNotification(from.displayName, from.emailAddress);
            });
        }
    }
})();