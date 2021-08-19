<!-- markdownlint-disable MD002 MD041 -->

This tutorial teaches you how to build an Office Add-in for Excel that uses the Microsoft Graph API to retrieve calendar information for a user.

> [!TIP]
> If you prefer to just download the completed tutorial, you can download or clone the [GitHub repository](https://github.com/microsoftgraph/msgraph-training-office-addin).

## Prerequisites

Before you start this demo, you should have [Node.js](https://nodejs.org) and [Yarn](https://yarnpkg.com/) installed on your development machine. If you do not have Node.js or Yarn, visit the previous link for download options.

> [!NOTE]
> Windows users may need to install Python and Visual Studio Build Tools to support NPM modules that need to be compiled from C/C++. The Node.js installer on Windows gives an option to automatically install these tools. Alternatively, you can follow instructions at [https://github.com/nodejs/node-gyp#on-windows](https://github.com/nodejs/node-gyp#on-windows).

You should also have either a personal Microsoft account with a mailbox on Outlook.com, or a Microsoft work or school account. If you don't have a Microsoft account, there are a couple of options to get a free account:

- You can [sign up for a new personal Microsoft account](https://signup.live.com/signup?wa=wsignin1.0&rpsnv=12&ct=1454618383&rver=6.4.6456.0&wp=MBI_SSL_SHARED&wreply=https://mail.live.com/default.aspx&id=64855&cbcxt=mai&bk=1454618383&uiflavor=web&uaid=b213a65b4fdc484382b6622b3ecaa547&mkt=E-US&lc=1033&lic=1).
- You can [sign up for the Office 365 Developer Program](https://developer.microsoft.com/office/dev-program) to get a free Office 365 subscription.

> [!NOTE]
> This tutorial was written with Node version 14.15.0 and Yarn version 1.22.0. The steps in this guide may work with other versions, but that has not been tested.

## Feedback

Please provide any feedback on this tutorial in the [GitHub repository](https://github.com/microsoftgraph/msgraph-training-office-addin).
