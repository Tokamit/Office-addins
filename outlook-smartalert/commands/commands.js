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
const i18n = {
    'ko-kr' : {
        faildToCheck : '수신자 확인을 실패하였습니다. =>',
        sendToExternal : '이 메일은 사외로 송신됩니다.',
    },
    'en-us' : {
        aildToCheck : 'Failed to Check',
        sendToExternal : 'This mail send to External Domain',
    },
    'ja-jp' : {
        aildToCheck : 'Failed to Check',
        sendToExternal : 'This mail send to External Domain',
    }
};


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
        recipients['bcc'] = [];
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
    console.log(Office.context.displayLanguage);
    console.log(i18n[Office.context.displayLanguage]);

    
    let udomains = [...new Set(domains)];
    let diff = udomains.filter(x => !TOKGRDOMAINS.includes(x));
    if (diff.length > 0){
        event.completed({ allowEvent: false, errorMessage: `test\n${diff.join("\n")}`,});
    } else {
        event.completed({ allowEvent: true});
    }
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
