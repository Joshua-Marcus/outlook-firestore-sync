# Outlook Add-In: Firestore Sync

<!-- [START badges] -->
[![License](https://img.shields.io/npm/l/ngx-auto-table.svg)](https://github.com/Joshua-Marcus/outlook-firestore-sync/blob/master/LICENSE) 
<!-- [END badges] -->

[![paypal](https://www.paypalobjects.com/en_US/i/btn/btn_donateCC_LG.gif)](https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=joshuamarcuswebdesign%40gmail.com&item_name=Development&currency_code=AUD&source=url)

A simple Outlook Add-In to save emails to Google Firestore.

### Setup

Create a config.ts file in the src/env directory

Config needs to export an object that contains emailCollectionPath, attachmentStoragePath and firebaseConfig

Example

```
export const config = {
  emailCollectionPath: 'PATH TO SAVE EMAILS',
  attachmentStoragePath: 'PATH TO SAVE ATTACHMENTS',
  firebaseConfig: {
    apiKey: "",
    authDomain: "",
    databaseURL: "",
    projectId: "",
    storageBucket: "",
    messagingSenderId: "",
    appId: ""
  }
} 

```


#### Running Locally
`npm run dev-server`


#### Updating Assets

To change the logo used replace the files found under assets/
