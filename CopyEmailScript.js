$("#get-email").click(getEmail);

function getEmail() {
  Office.context.mailbox.item.body.getAsync("text", processBody);
  function processBody(res) {
    var msgSender = Office.context.mailbox.item.sender;
    var msgTo = Office.context.mailbox.item.to;
    var msgCc = Office.context.mailbox.item.cc;
    var msgSubject = Office.context.mailbox.item.subject;
    var msgDate = Office.context.mailbox.item.dateTimeModified.toString();

    var nm = msgSender.displayName.split(" ")[0];
    var words = res.value.trim().split("From:")[0];

    if (words.includes(nm)) words = words.split(nm)[0] + "\r\n" + msgSender.displayName;

    var mto = new Array();
    for (var k = 0; k < msgTo.length; k++) {
      mto[k] = msgTo[k].emailAddress;
    }

    var ccd = new Array();
    for (var i = 0; i < msgCc.length; i++) {
      ccd[i] = msgCc[i].emailAddress;
    }

    var input = document.createElement("textarea");
    input.innerHTML =
      "FROM: " +
      msgSender.displayName +
      " [" +
      msgSender.emailAddress +
      "]" +
      "\r\n" +
      "TO: " +
      mto +
      "\r\n" +
      "CC: " +
      ccd +
      "\r\n" +
      "SUBJECT: " +
      msgSubject +
      "\r\n" +
      "DATE: " +
      msgDate
        .split(/\s+/)
        .slice(1, 4)
        .join(" ") +
      "\r\n" +
      "\r\n" +
      words.replace(/\n+/g, "\n");
    document.body.appendChild(input);
    input.select();
    var result = document.execCommand("copy");

    document.body.removeChild(input);
  }
}
