import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
import firebase from "firebase/app";
import "firebase/auth";
import { firebaseConfig } from "../env/config";

/* global document, Office */

var authRecord: firebase.auth.UserCredential;

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    firebase.initializeApp(firebaseConfig);
    document.getElementById("login-btn").onclick = login;
    document.getElementById("logout-btn").onclick = logout;
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