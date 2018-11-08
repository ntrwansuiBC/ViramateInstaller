console.log("Started");

chrome.runtime.onMessage.addListener(onExternalMessage);
chrome.runtime.onMessageExternal.addListener(onExternalMessage);

function log (...args) {
    args.unshift((new Date()).toLocaleTimeString("en-US", {hour12: false}) + " |");
    console.log.apply(console, args);
};

function onExternalMessage (message, sender, sendResponse) {
    if (!message || !message.type)
        return;
    var type = message.type;

    log("Processing external message", type);

    switch (type) {
        case "reloadExtension":
            reloadExtension(message.id, message.reason);
            break;
    }
};

function reloadExtension (id, reason) {
    // FIXME: Reloading an extension this way does not update the version number and other manifest info.

    log("Reloading extension", id, reason);
    chrome.management.setEnabled(id, false, function () {
        log("Extension disabled");
        chrome.management.setEnabled(id, true, function () {
            log("Extension re-enabled. Reload complete.");
        });
    });
};