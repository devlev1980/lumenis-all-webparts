## external-sharing

## Summary
This Development shows how to create Custom Dialogs using `@microsoft/sp-dialog` package in the context of Command View Set and send selected document with e-mail.

## Solution

Solution|Author(s)
--------|---------
Extarnal-Sharing | PzProjects

## Version history

Version|Date|Comments
-------|----|--------
1.0|July 30, 2020|Initial version

## Technology versions used

* Node.js- v8.11.4
* Gulp-
  CLI version: 2.2.0
  Local version: 3.9.1
* Npm- 6.12.0

## Set your environment

Please follow this guide in order to set up your SharePoint Framework development environment:
[SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment)

## Preparations
- In case of general changes please make sure that Constants (External-Sharing\src\extensions\externalSharing\models\Constants.ts) 
  Is updated and the relevent variabiles didn't change.
  For exampale: "ARRAY_OF_ACTIVE_LIBRARY_NAMES" which indicates the librarys that uses this development.
- Every document library that uses this development must contain a text column name "ExternalSharingStatus" (you can change the name after creating the column)
  this column indicates that the file status.

## Debug

- First you need to set the url of the required document library:
  Go to "serve.json" under the folder "config" and change "pageUrl" to the required document library url
  This step is not mandatory when the relevent document library remain the same!
- Move to folder where this readme exists
- In the command line run:
  - `npm install`
  - `gulp trust-dev-cert`
  - `gulp serve`

## Deploy

- Move to folder where this readme exists
- In the command line run:
  - `gulp serve --nobrowser`
  - `gulp clean`
  - `gulp bundle --ship`
  - `gulp package-solution --ship`
- Upload .sppkg file from sharepoint\solution to your tenant App Catalog
  E.g.: https://<tenant>.sharepoint.com/sites/AppCatalog/AppCatalog
- Only on the first upload: you need to approve Graph API request in the office365 admin center:
  office365 admin >>> Advanced (right pannel) >>> Api access 

## Features

This project contains SharePoint Framework extensions that illustrates next features:
* Command extension
  [ListView Command Set](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/extensions/get-started/building-simple-cmdset-with-dialog-api)
* Custom dialog control using `@microsoft/sp-dialog` package
* Using @pnp/sp
* Using Microsoft Graph API
  [Consume Microsoft Graph](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/use-aad-tutorial)
* React
* Using `@fluentui/react` A collection of UX frameworks
  Fluent UI React is the official open-source React front-end framework designed to build experiences that fit seamlessly into a broad range of Microsoft products.
  [fluentui](https://developer.microsoft.com/en-us/fluentui#/get-started)