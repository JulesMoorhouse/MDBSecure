MDBSecure
=========

![GitHub](https://img.shields.io/github/license/jules2010/ideaspad.svg) ![GitHub contributors](https://img.shields.io/github/contributors/jules2010/ideaspad.svg) ![GitHub issues](https://img.shields.io/github/issues/jules2010/ideaspad.svg) ![GitHub pull requests](https://img.shields.io/github/issues-pr/jules2010/ideaspad.svg)
[![Languages](https://img.shields.io/badge/language-vb.net-FF69B4.svg)](#)

### Is YOUR MICROSOFT ACCESS DATABASE Secure?

Microsoft Access databases are not secure by default. Database security features
are included in Microsoft Access but different versions of the database have
different security weaknesses. 

There is no single method of making a secure Access database. If you have the
time, Microsoft provide instructions for manually encrypting passwords and
enabling security. The security-enabling process can take half an hour per
database even if you are technically knowledgeable.

Our database security software, MDBSecure, makes the process of securing a
database incredibly easy. Using our security utility, you never be put off
making your databases secure again. Our security software will aid you in
closing any security gaps or loopholes in your organisation's data storage.

If you want to find out more about how our security software can help you
protect your data, read on below.

Written by [Jules Moorhouse](https://www.julesmoorhouse.com).

# Table of contents
* [Background](#background)
* [Features and functionality](#features-and-functionality)
* [How to build / edit the code](#how-to-build--edit-the-code)
* [Installation](#installation)

# Features and functionality

### What is MDBSecure?
* An MS Access Security utility, providing protection for your databases.
* MDBSecure follows the recommended approach to securing Microsoft&reg; Access databases, it will encrypt your access database and remove security weaknesses.
* SecFaq - Security FAQ (Recommended approach by Microsoft)
* Although nothing is ever 100% secure in computing, following this approach will make an MS Access Database more secure.
* For more information about the theory behind securing MS Access databases please see the [SECFAQ](http://web.archive.org/web/20080424151236/http://support.microsoft.com/kb/165009) from the internet archive. SECFAQ which is shortened from 'security FAQ' document.

* For more technical information on how MDBSecure works, download my [Technical Database Security Information White Paper.](https://github.com/Jules2010/MDBSecure/blob/master/whitepaper.pdf)

<br />

## Features                                      

|||
| :- | - |
| Removes Admin user permissions | Append tables from another DB |
| Assigns new super user | Creates Jet Connection Strings |
| Encrypt database | Opens Secure Databases with MS Access |
| MS Access 2000 / 2002 / 2003  format | Check for Updates |
| Workgroup creation | Easy to use, help provided |

# Background

The functionality for MDBSecure was initially written as part of Ideaspad in 2002. I wanted try and make the data used by Ideaspad as secure as possible.

I didn’t want users to modify the data / structure and cause issues with Ideaspad. 

Obviously MS Access isn’t bulletproof, however my aim was to make it as secure as possible.

I spent a lot of time researching MS Access security and figuring out the steps required.

Later I put all this functionality into the MDBSecure application.

# How to build / edit the code

You’ll need Microsoft Visual Basic .Net 2003, you can download this as part of your MSDN benefits.

You will also probably need to use Windows 7 or earlier.

Perhaps in the future I may upgrade this project if this becomes a popular request.

You’ll also need to install Microsoft Data Access Components 2.8.

I have included a Components folder which includes most of the required DLLs.

You’ll notice there are two project files for Ideaspad and a number of DLLs suffixed with Debug, these are added to one solution for easier debugging. (Why, well there needs to be separate solution to allow for obfuscating / strong name signing, which is used in the batch files.)

So the easiest way to compile the project is to open \CodeLibrary\SharewareProjs\MDBSecure\MDBSecureDebug.sln

However you can also use the batch files in the build folder, start with the DLLs.

<br/>

# Installation

You can install the app via [MDBSecure.msi](https://github.com/Jules2010/MDBSecure/blob/master/Build/MDBSecure/DistBuild/MDBSecure.msi) which will inform you about any system requirements you don't already have.

<br/>

# Contributing
Contributions for bug fixing or improvements are welcomed. Feel free to submit a pull request.

<br/>

# License
Usage is provided under the [MIT License](http://opensource.org/licenses/mit-license.php). See LICENSE for the full details.

