import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
import firebase from "firebase/app";
import "firebase/auth";
import { firebaseConfig } from "../env/config";

/* global document, Office */

var authRecord: firebase.auth.UserCredential;

Office.onReady(info => {
  // Add Event listener on email change - check if current email is already synced.
  
  if (info.host === Office.HostType.Outlook) {
    firebase.initializeApp(firebaseConfig);
    document.getElementById("login-btn").onclick = login;
    document.getElementById("logout-btn").onclick = logout;
    document.getElementById("sync-btn").onclick = syncEmail;
    initAddInAuth();

  }
});

function initAddInAuth() {

  const isLoggedIn = checkAuth();
  if(isLoggedIn) {
    document.getElementById("auth").style.display = "unset";
    document.getElementById("no-auth").style.display = "none";
    document.getElementById("authEmail").innerHTML = authRecord.user.email;
  } else {
    document.getElementById("no-auth").style.display = "unset";
  }
}


async function login() {
  try {
    await firebase.auth().setPersistence(firebase.auth.Auth.Persistence.LOCAL);
    var formEmail = (<HTMLInputElement>document.getElementById("email")).value;
    var formPassword = (<HTMLInputElement>document.getElementById("password")).value;
    const result = await firebase.auth().signInWithEmailAndPassword(formEmail, formPassword);
    localStorage.setItem('authResult', JSON.stringify(result));
    initAddInAuth();
  } catch (e) {
    document.getElementById("login-result").innerHTML = e.message;
    document.getElementById("login-result").style.color = "#f44336";
    console.log(e);

  }
}


async function syncEmail() {
  try {
    console.log('Running actions')
    const userID = firebase.auth().currentUser ? 'yes' : 'no';
    console.log(userID)
           const noAuthMessage: Office.NotificationMessageDetails = {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: `${userID}`,
      icon: "Icon.80x80",
      persistent: true
    };
    Office.context.mailbox.item.notificationMessages.replaceAsync("action", noAuthMessage);
  //   console.log('isUserLoggedIn', isLoggedIn);
  //   if(!isLoggedIn) {
  //       const noAuthMessage: Office.NotificationMessageDetails = {
  //     type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
  //     message: "Please login to use this function.",
  //     icon: "Icon.80x80",
  //     persistent: false
  //   };
  //   Office.context.mailbox.item.notificationMessages.replaceAsync("action", noAuthMessage);
  //   }

    var bodyResult;

    const mailRecord = Office.context.mailbox.item;
    // mailRecord.getAttachmentsAsync(async function(attachment) {
    //   attachment.value.map(a => a.url)
    // })
    // const attachemnets = mailRecord.attachments;
    // attachemnets.map(at => at)
    const itemId = mailRecord.itemId
    
    var body = Office.context.mailbox.item.body;
    body.getAsync(Office.CoercionType.Html, async function (asyncResult) {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          bodyResult = "";
      } else {
          bodyResult = `${asyncResult.value}`;
          await firebase.firestore().collection('users').doc(authRecord.user.uid).collection('synced-emails').doc(itemId).set({
            from: {name: mailRecord.from.displayName, email: mailRecord.from.emailAddress},
            userId: authRecord.user.uid,
            subject: mailRecord.subject,
            body: bodyResult,
            id: itemId
          }, {merge: true});
          const completed: Office.NotificationMessageDetails = {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: "You email has been synced!",
            icon: "Icon.80x80",
            persistent: false
          };
          Office.context.mailbox.item.notificationMessages.replaceAsync("action", completed);

      }
   });

  } catch (e) {
    // console.log(e);
    // event.completed();

  }

}

async function logout() {
  await firebase.auth().signOut();
  localStorage.clear();
  document.getElementById("auth").style.display = "none";
  document.getElementById("no-auth").style.display = "unset";
}


function checkAuth(): boolean {
  if(localStorage.getItem('authResult')) {
    authRecord = JSON.parse(localStorage.getItem('authResult'));
    return true;
  } else {
    localStorage.clear();
    return false;
  }

}