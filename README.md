MDBSecure
=========

![GitHub](https://img.shields.io/github/license/jules2010/ideaspad.svg) ![GitHub contributors](https://img.shields.io/github/contributors/jules2010/ideaspad.svg) ![GitHub issues](https://img.shields.io/github/issues/jules2010/ideaspad.svg) ![GitHub pull requests](https://img.shields.io/github/issues-pr/jules2010/ideaspad.svg)
[![Languages](https://img.shields.io/badge/language-vb.net-FF69B4.svg)](#)

XXXXXX

Written by [Jules Moorhouse](https://www.julesmoorhouse.com).

# Table of contents
* [Background](#background)
* [Features and functionality](#features-and-functionality)
* [How to build / edit the code](#how-to-build--edit-the-code)
* [Installation](#installation)

# Features and functionality

XXXXX

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

