# Getting PowerShell Support

**This repository is NOT a support platform!**
Its intent is to provide information and tools to aid you in getting better help from whoever you contact for help with PowerShell.

## Official PowerShell Team Support

To receive formal, official support in a PowerShell matter, either open a case with Microsoft Support or [Open an Issue on Github](https://github.com/PowerShell/PowerShell/issues).

Anything you see below can help you provide better information, but this project is not the place to open issues for any problems you have.

## Community Support

There are plenty of options to get help from the public community.
[Find them listed here.](https://learn.microsoft.com/en-us/powershell/scripting/community/community-support)

## Information Collection Scripts

While we go into more details further down in this readme, odds are you are just here to run that support data gathering script, so here you go:

> Support Package

For when you want one debug dump for an export to look at. Contains the most information, but requires file transfer.

```powershell
Invoke-Expression ([System.Net.WebClient]::new()).DownloadString('https://raw.githubusercontent.com/microsoft/powershell-getsupport/main/support-package.ps1')
```

> Support Data

Provides a comprehensive debug summary as a text file. Less information than the support package, but can be used without sending files.

```powershell
Invoke-Expression ([System.Net.WebClient]::new()).DownloadString('https://raw.githubusercontent.com/microsoft/powershell-getsupport/main/support-message.ps1')
```

## Do a clean run

If at all possible, ...

+ Start a new console
+ Reproduce the issue you want support with
+ Collect Debug information

Collecting data from a console you have been working with for a day or two adds a lot of distractions that make it harder to pinpoint the problem.
Keep the problem reproduction as simple as possible.

## What Information to Provide

There is a lot of information on any given computer, but what information do you really need?

> Scale / Scope

Does this happen on a single machine or on many? All?
Does it apply only when running against one specific target but does not affect others (if any)?
If it happens on many but not all machines, what do they have in common?

Clarifying this information - which a debug dump cannot collect for you - is really helpful both with prioritizing an issue but also with where to start looking for the issue.

> Runtime Environment (include in the scripts above)

+ PS Version
+ Operating System
+ Behind a Proxy?
+ Modules and versions loaded
+ PSSnapins and versions loaded
+ Assemblies, versions and paths loaded

> What happened?

+ Command History (include in the scripts above)
+ Code executed
+ Error data (include in the scripts above)
