/*
 * Copyright (c) TOKAM
 * Outlook Add-in External Domain Alert
 * v1.0 initialize
 * v1.1 i18n
 * v1.2 normal to arrow func
 */

Office.onReady();

/**
 * whitelist domains
 */
const WLDOMAINS = [
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
 * i18n
 * base on Office.context.displayLanguage
 * Refer https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/6c085406-a698-4e12-9d4d-c3b0ee3dbc4a
 */
const i18n = {
    'ko-KR' : {
        faildToCheck : '수신자 확인을 실패하였습니다. =>',
        sendToExternal : '이 메일은 사외로 송신됩니다. 다음의 외부도메인을 확인해주세요',
    },
    'en-US' : {
        faildToCheck : 'Failed to Check Recipients. =>',
        sendToExternal : 'This mail send to External Domain. Please check to follow external domains.',
    },
    'ja-JP' : {
        faildToCheck : '宛先確認に失敗しました。=>',
        sendToExternal : '本メールは社外へ送信されます。次の外部ドメインを確認してください。',
    }
};

let l10n = null;

/**
 * create new mail or appointment
 */
function onItemComposeHandler(event) {
    setl10n();
    event.completed({ allowEvent: true });
}

/**
 * send to mail or appointment
 */
function onItemSendHandler(event) {
    event.completed({ allowEvent: false, errorMessage:"alert test." });
}

/**
 */
function setl10n(){
    l10n = i18n[Office.context.displayLanguage] ?? i18n['en-US']
}

/**
 */
function diplayMessageBoxExternalDomain(event,domains){
    let diff = [...new Set(domains)].filter(e => !WLDOMAINS.includes(e));
    let param = diff.length > 0 ? { allowEvent: false, errorMessage: `${l10n.sendToExternal}\n${diff.join("\n")}`,} : { allowEvent: true};
    event.completed(param);
}

/**
 */
function getRecipiensDomain(recipients){
    return recipients.map(e => e.emailAddress.split('@').pop());
}

/**
 * event handler bind for outlook application
 */
Office.actions.associate("onMessageComposeHandler", onItemComposeHandler);
Office.actions.associate("onAppointmentComposeHandler", onItemComposeHandler);
Office.actions.associate("onMessageSendHandler", onItemSendHandler);
Office.actions.associate("onAppointmentSendHandler", onItemSendHandler);
