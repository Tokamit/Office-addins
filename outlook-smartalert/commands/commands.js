/*
 * Copyright (c) TOKAM
 */

Office.onReady();

/**
 */

const TOKGRDOMAINS = [
    "tok.co.jp", 
    "tokamerica.com", 
    "tokam.co.kr", 
    "toktaiwan.com.tw", 
    "tokchina.com.cn", 
    "ohka.nl", 
    "tok.com.sg", 
    "tokgroup.onmicrosoft.com"
];


/**
 */
function onItemComposeHandler(event) {
    console.log("email compose");
    event.completed({ allowEvent: true });
}

/**
 */
function onItemSendHandler(event) {
    let toRecipients, ccRecipients, bccRecipients;
    let item = Office.context.mailbox.item;

    if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    } else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }

    toRecipients.getAsync(
        { asyncContext: { callingEvent: event } },
        (asyncResult) => {
            let event = asyncResult.asyncContext.callingEvent;
            let domains =[];
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                event.completed({ allowEvent: false, errorMessage: "Failed to configure categories.",});
                return;
            }
            displayAddresses(asyncResult.value);
            domains.push(getRecipiensDomain(asyncResult.value));
            console.log(domains)
    });

    ccRecipients.getAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            write(asyncResult.error.message);
            return;
        }
        displayAddresses(asyncResult.value);
    });

    if (bccRecipients.length > 0) {
        bccRecipients.getAsync((asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                write(asyncResult.error.message);
                return;
            }
            displayAddresses(asyncResult.value);
        });
    } else {
        write("Recipients in the Bcc field: None");
    }

}

function getRecipiensDomain(recipients){
    let values = [];
    recipients.forEach((recipient) => {
        values.push(recipient.emailAddress);
    });
    return values;
}

function displayAddresses (recipients) {
    recipients.forEach((recipient) => {
        console.log(recipient.emailAddress);
    });
}



Office.actions.associate("onMessageComposeHandler", onItemComposeHandler);
Office.actions.associate("onAppointmentComposeHandler", onItemComposeHandler);
Office.actions.associate("onMessageSendHandler", onItemSendHandler);
Office.actions.associate("onAppointmentSendHandler", onItemSendHandler);
