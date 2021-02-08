import firebase from "firebase/app";
import "firebase/firestore";
import { firebaseConfig } from "../env/config";
import "firebase/auth";

var authRecord: firebase.auth.UserCredential;
function checkAuth(): boolean {
  if(localStorage.getItem('authResult')) {
    authRecord = JSON.parse(localStorage.getItem('authResult'));
    return true;
  } else {
    localStorage.clear();
    return false;
  }

}

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
  firebase.initializeApp(firebaseConfig);
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
async function action(event: Office.AddinCommands.Event) {
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
  //   var bodyResult;

  //   const mailRecord = Office.context.mailbox.item;
    
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
  //           body: bodyResult
  //         });
  //         const completed: Office.NotificationMessageDetails = {
  //           type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
  //           message: "You email has been synced!",
  //           icon: "Icon.80x80",
  //           persistent: false
  //         };
  //         Office.context.mailbox.item.notificationMessages.replaceAsync("action", completed);

  //     }
  //  });

    event.completed();
  } catch (e) {
    console.log(e);
    event.completed();

  }

}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal() as any;

// the add-in command functions need to be available in global scope
g.action = action;

