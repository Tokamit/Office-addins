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
    let recipients = {};
    let item = Office.context.mailbox.item;

    if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
        recipients['to'] = item.requiredAttendees;
        recipients['cc'] = item.optionalAttendees;
    } else {
        recipients['to'] = item.to;
        recipients['cc'] = item.cc;
        recipients['bcc'] = item.bcc;
    }

    recipients['to'].getAsync(
    { asyncContext: { callingEvent: event, recipients: recipients } },
    (asyncResult) => {
        let event = asyncResult.asyncContext.callingEvent;
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            event.completed({ allowEvent: false, errorMessage: "Failed to Check To",});
            return;
        }
        let domains =[];

        displayAddresses(asyncResult.value);//STUFF
        getRecipiensDomain(asyncResult.value).forEach(e=>{domains.push(e)});
        
        recipients['cc'].getAsync(
        { asyncContext: { callingEvent: event, recipients: recipients, domains:domains } },
        (asyncResult) => {
            let event = asyncResult.asyncContext.callingEvent;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                event.completed({ allowEvent: false, errorMessage: "Failed to Check CC",});
                return;
            }
            displayAddresses(asyncResult.value);//STUFF
            getRecipiensDomain(asyncResult.value).forEach(e=>{domains.push(e)});

            if (recipients['bcc'].length > 0) {
                recipients['bcc'].getAsync(
                { asyncContext: { callingEvent: event, recipients: recipients, domains:domains } },
                (asyncResult) => {
                    let event = asyncResult.asyncContext.callingEvent;
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        event.completed({ allowEvent: false, errorMessage: "Failed to Check BCC",});
                        return;
                    }
                    displayAddresses(asyncResult.value);//STUFF
                    getRecipiensDomain(asyncResult.value).forEach(e=>{domains.push(e)});
                    diplayMessageBoxExternalDomain(event,domains);
                });
            } else {
                diplayMessageBoxExternalDomain(event,domains);
            }
        });
    });

    

}

function diplayMessageBoxExternalDomain(event,domains){
    let udomains = [...new Set(domains)];
    console.log(udomains);
    //check domain
    event.completed({ allowEvent: false, errorMessage: "This mail send to External Domain",});
}

function getRecipiensDomain(recipients){
    let values = [];
    recipients.forEach((recipient) => {
        values.push(recipient.emailAddress.split('@').pop());
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
