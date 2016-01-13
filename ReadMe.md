## Synopsis

The purpose of this PowerShell-MassInstallPackage.ps1 script is to enable an admin to assign a silent install package to one 
or more computers that are a part of an existing PSMANAGE scripting tool environment. 

For more information about the PSMANAGE scripts, check out: https://gregbesso.wordpress.com/projects/psmanage/
For more information about the LaunchPad script, check out: https://gregbesso.wordpress.com/projects/powershell-launchpad-script/

The script reads from the PSMANAGE SharePoint lists to gather information about available software installation packages, and 
also to get a list of computers that tasks can be assigned to. The admin picks a package, picks one or more computers, and 
then clicks a button and tasks are created.

Then on their own interval, the computers check in to get their tasks and then perform each one. Once the tasks are performed, 
the computers will update their tasks with details of the results. That is part of another script set, the PSMANAGE scripts.


## Code Example

The LaunchPad lets the admin select this scripting tool from their list of tools, if being used. Otherwise the admin can run this 
script directly.


## Motivation

I wanted to have a better way of deploying silent install packages to computers. In larger deployments, not all computers were online 
at the same time and I didn't want to have to go follow up / look for when they were back online. I wanted to build my own "home rolled" 
SMS / SCCM solution that tapped into the PSMANAGE scripting tool that I already had in place. The information was there, the phoning 
home was there, so I figured it would be easy to add this useful functionality.



## Installation

This PowerShell-MassInstallPackage.ps1 file is meant to be run using the LaunchPad parent script, but can also be run stand-alone.
There is a required step, even after you have the PSMANAGE scripts up and running. The steps you should follow...

1) Get your PSMANAGE scripts installed and running. Without that, this won't work.
2) In the PSMANAGE-CentralServer-Imports.ps1 file, go to the Get-PSManageCentralServerStarted() function and uncomment 
the New-PSManageListsPackagesTasks line.
3) Wait for your server to run on the next scheduled time for your job, or run it manually.
4) You then need to access your new empty PSMANAGE-Packages list and enter one or more silent install packages into the list.
5) Then at that point you can use this script to create new tasks!
6) PS you can also just directly create tasks in your new empty PSMANAGE-Tasks list too, but this script makes it much easier 
to make tasks for many computers seamlessly.


## API Reference

No API here. <sounds of crickets>

## Tests

No testing info here. <sounds of crickets>

## Contributors

Just a solo script project by moi, Greg Besso. Hi there :-)

## License

Copyright (c) 2015 Greg Besso

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.