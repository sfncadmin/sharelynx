/* global Office */
// OnMessageSend Smart Alert: warn when large attachments are about to be sent.
// Runs in classic Outlook's JavaScript-only runtime, so this file stays ES5:
// no imports, no DOM, no arrow functions, no template literals.

var DEFAULT_THRESHOLD_MB = 5;

function getThresholdMB() {
  try {
    var rs = Office.context.roamingSettings;
    var v = rs && rs.get("lynx_thresholdMB");
    if (typeof v === "number" && v >= 0) return v;
    if (typeof v === "string" && v !== "" && !isNaN(Number(v))) return Number(v);
  } catch (e) {
    // settings unavailable in this runtime; fall through to default
  }
  return DEFAULT_THRESHOLD_MB;
}

function formatMB(bytes) {
  return (Math.round((bytes / (1024 * 1024)) * 10) / 10).toString();
}

function onMessageSendHandler(event) {
  try {
    var thresholdMB = getThresholdMB();
    if (thresholdMB === 0) {
      event.completed({ allowEvent: true });
      return;
    }
    Office.context.mailbox.item.getAttachmentsAsync({ asyncContext: event }, function (result) {
      var ev = result.asyncContext;
      try {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          ev.completed({ allowEvent: true }); // fail open - never block send on our own error
          return;
        }
        var limit = thresholdMB * 1024 * 1024;
        var big = [];
        for (var i = 0; i < result.value.length; i++) {
          var a = result.value[i];
          if (a.attachmentType === "file" && !a.isInline && a.size > limit) {
            big.push(a.name + " (" + formatMB(a.size) + " MB)");
          }
        }
        if (big.length === 0) {
          ev.completed({ allowEvent: true });
          return;
        }
        ev.completed({
          allowEvent: false,
          errorMessage:
            "This message has " + big.length + " attachment(s) over " + thresholdMB + " MB:\n" +
            big.join("\n") +
            "\n\nOpen ShareLynx to upload them to OneDrive and send a link instead, or send anyway.",
          cancelLabel: "Convert to links",
          commandId: "msgComposeOpenPaneButton"
        });
      } catch (inner) {
        ev.completed({ allowEvent: true });
      }
    });
  } catch (outer) {
    event.completed({ allowEvent: true });
  }
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
