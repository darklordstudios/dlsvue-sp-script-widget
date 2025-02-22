# dlsvue-sp-script-widget

## Summary

Used to load external scripts built with other frameworks such as VueJS.

SharePoint web part context is added to the globalThis object to pass to loaded script.

[picture of the solution in action, if possible]

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.20.0-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

> @pnp PowerShell

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| folder name | Author details (name, company, twitter alias with link) |

## Version history

| Version | Date                | Comments        |
| ------- | ------------------- | --------------- |
| 1.1     | February 22, 2025   | Update to globalThis  |
| 1.0     | November, 2024      | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- Ensure you have installed the prerequisites for SharePoint Framework development. [Installed gulp, etc]
- in the command-line run:
  - **npm install**
  - **Update the make.ps1 and install.ps1 files for your install path and SharePoint site**
  - **.\make.ps1**
  - **.\install.ps1**

> Include any additional steps as needed.

## Features

This webpart solution creates a web part that will allow you select a script to load into the web part.

You can create the scripts how you like but loading them to the webpart requires a few steps in SharePoint.

Your script files should be loaded into a document library. For my examples and the way I recommend doing it is to create folders [Yes, an exception] in the Site Assets library.
Create HTML, CSS, and JS folders. Add your files to the respective folders.
It is recommended to use .txt files for the main script code and to create this file to load scripts and css through code. An example solution can be found at: 

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
