/*
 * Copyright (c) TOKAM
 */

Office.onReady();

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

let _recipients = [];

/**
 */
function onItemComposeHandler(event) {
    console.log("event compose");
    event.completed({ allowEvent: true });
}

/**
 */
function onItemSendHandler(event) {
    let item, itemDomains;
    item = Office.context.mailbox.item;

    let toRecipients, ccRecipients, bccRecipients;
    let recipients;

    if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    } else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }

    toRecipients.getAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            write(asyncResult.error.message);
            return;
        }
        _recipients += asyncResult.value;
    });

    displayAddresses(_recipients);

    event.completed({
        allowEvent: false,
        errorMessage: "test",
      });

}


function displayAddresses (recipients) {
    for (let i = 0; i < recipients.length; i++) {
        write(recipients[i].emailAddress);
    }
}
function write(message) {
    console.log(message);
}
