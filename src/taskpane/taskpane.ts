import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
import firebase from "firebase/app";
import "firebase/auth";
import "firebase/storage";
import { config } from "../env/config";
import * as md5 from "md5";

interface Attachment {
  contentType: string;
  name: string;
  content: string;
  downloadURL?: string;
}

/* global document, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    firebase.initializeApp(config.firebaseConfig);
    firebase.auth().onAuthStateChanged(user => {
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
    if (!isLoggedIn) {
      const noAuthMessage: Office.NotificationMessageDetails = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Please login to use this function.",
        icon: "Icon.80x80",
        persistent: false
      };
      Office.context.mailbox.item.notificationMessages.replaceAsync("action", noAuthMessage);
    }
    const mailRecord = Office.context.mailbox.item;

    const storageRef = firebase.storage().ref();

    const idString: string = md5(mailRecord.itemId);
    const body = await getMailBody(Office.CoercionType.Html);
    let attachments = await getMailAttachments();
    attachments = await Promise.all(
      attachments.map(async attachment => {
        const imageRef = storageRef.child(config.attachmentStoragePath + attachment.name);
        const storageObj = await imageRef.putString(
          `data:${attachment.contentType};base64, ` + attachment.content,
          "data_url"
        );
        const downloadURL = await storageObj.ref.getDownloadURL();
        attachment.downloadURL = downloadURL;
        return attachment;
      })
    );
    const dbMailObj = {
      id: idString,
      body,
      subject: mailRecord.subject,
      from: { name: mailRecord.from.displayName, email: mailRecord.from.emailAddress },
      attachments: attachments,
      userId: isLoggedIn.uid
    };
    console.log({ dbMailObj });
    await firebase
      .firestore()
      .collection(config.emailCollectionPath)
      .add(dbMailObj);
    const completed: Office.NotificationMessageDetails = {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "Your email has been synced!",
      icon: "Icon.80x80",
      persistent: false
    };
    Office.context.mailbox.item.notificationMessages.replaceAsync("action", completed);
  } catch (e) {
    console.log(e);
  }
}

async function getMailBody(coercionType: Office.CoercionType) {
  return new Promise((resolve, reject) => {
    const body = Office.context.mailbox.item.body;
    body.getAsync(coercionType, function(asyncResult) {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reject();
      } else {
        resolve(asyncResult.value);
      }
    });
  });
}

async function getMailAttachments(): Promise<Attachment[]> {
  const attachments = Office.context.mailbox.item.attachments;
  const attachmentsNotInline = attachments.filter(attachment => !attachment.isInline);
  const attachmentContent = await Promise.all(
    attachmentsNotInline.map(async attachment => {
      return await new Promise((resolve, reject) => {
        Office.context.mailbox.item.getAttachmentContentAsync(attachment.id, function(asyncResult) {
          if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            reject();
          } else {
            resolve({
              content: asyncResult.value.content,
              contentType: attachment.contentType,
              name: attachment.name
            } as Attachment);
          }
        });
      });
    })
  );

  return attachmentContent as Attachment[];
}

async function logout() {
  await firebase.auth().signOut();
  document.getElementById("auth").style.display = "none";
  document.getElementById("no-auth").style.display = "unset";
}
