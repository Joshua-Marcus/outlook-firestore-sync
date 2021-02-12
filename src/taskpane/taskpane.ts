import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
import firebase from "firebase/app";
import "firebase/auth";
import { firebaseConfig } from "../env/config";
import * as md5 from 'md5'

/* global document, Office */


Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    firebase.initializeApp(firebaseConfig);
    firebase.auth().onAuthStateChanged((user) => {
      if (user != null) {
        document.getElementById("auth").style.display = "unset";
        document.getElementById("no-auth").style.display = "none";
        document.getElementById("authEmail").innerHTML = user.email;
      } else {
        document.getElementById("no-auth").style.display = "unset";
      }
    });
    document.getElementById("login-btn").onclick = login;
    document.getElementById("logout-btn").onclick = logout;
    document.getElementById("sync-btn").onclick = syncEmail;
  }
});




async function login() {
  try {
    await firebase.auth().setPersistence(firebase.auth.Auth.Persistence.LOCAL);
    var formEmail = (<HTMLInputElement>document.getElementById("email")).value;
    var formPassword = (<HTMLInputElement>document.getElementById("password")).value;
    await firebase.auth().signInWithEmailAndPassword(formEmail, formPassword);
  } catch (e) {
    document.getElementById("login-result").innerHTML = e.message;
    document.getElementById("login-result").style.color = "#f44336";
    console.log(e);

  }
}


async function syncEmail() {
  try {
    // Recheck Auth
    const isLoggedIn = firebase.auth().currentUser;
    if(!isLoggedIn) {
      const noAuthMessage: Office.NotificationMessageDetails = {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "Please login to use this function.",
      icon: "Icon.80x80",
      persistent: false
    };
    Office.context.mailbox.item.notificationMessages.replaceAsync("action", noAuthMessage);
    }

    // var bodyResult;

    const mailRecord = Office.context.mailbox.item;
    // mailRecord.getAttachmentsAsync(async function(attachment) {
    //   attachment.value.map(a => a.url)
    // })
    // const attachemnets = mailRecord.attachments;
    // attachemnets.map(at => at)
      const idString: string = md5(mailRecord.itemId);
      const body = await getMailBody(Office.CoercionType.Html);
      const attachments = await getMailAttachments();
      console.log(attachments);
      const dbMailObj = {
        id: idString,
        body,
        subject: mailRecord.subject,
        from: {name: mailRecord.from.displayName, email: mailRecord.from.emailAddress}
      }

      console.log(dbMailObj)
  //   var body = Office.context.mailbox.item.body;
  //   body.getAsync(Office.CoercionType.Html, async function (asyncResult) {
  //     if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
  //         bodyResult = "";
  //     } else {
  //         bodyResult = `${asyncResult.value}`;
  //         await firebase.firestore().collection('users').doc(authRecord.user.uid).collection('synced-emails').add({
  //           from: {name: mailRecord.from.displayName, email: mailRecord.from.emailAddress},
  //           userId: authRecord.user.uid,
  //           subject: mailRecord.subject,
  //           body: bodyResult,
  //           mailbox
  //           id: itemId
  //         }, {merge: true});
  //         const completed: Office.NotificationMessageDetails = {
  //           type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
  //           message: "You email has been synced!",
  //           icon: "Icon.80x80",
  //           persistent: false
  //         };
  //         Office.context.mailbox.item.notificationMessages.replaceAsync("action", completed);

  //     }
  //  });

  } catch (e) {
    console.log('Something went wrong', e)
    // console.log(e);
    // event.completed();

  }

}

async function getMailBody(coercionType: Office.CoercionType) {
  return new Promise((resolve, reject) => {
    const body = Office.context.mailbox.item.body;
    body.getAsync(coercionType, function (asyncResult) {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          reject()
      } else {
        resolve(asyncResult.value)

      }
   });
  })
}


async function getMailAttachments(): Promise<string> {
  return new Promise((resolve, reject) => {
    const attachments = Office.context.mailbox.item.attachments;
    console.log({attachments});
    Office.context.mailbox.item.getAttachmentContentAsync(attachments[0].id, function (asyncResult) {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reject();
      } else {
        console.log(asyncResult.value.content);
        resolve(asyncResult.value.content)
      }
    });
  });
}

async function logout() {
  await firebase.auth().signOut();
  document.getElementById("auth").style.display = "none";
  document.getElementById("no-auth").style.display = "unset";
}
