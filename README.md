# wine-prefix_O365
## Introduction
This repo contains Bash scripts for ease the installation of O365 via [Wine](https://wiki.winehq.org/Main_Page) and [Winetricks](https://github.com/Winetricks/winetricks) within Debian-based Linux distributions.

The scripts were initially made for my personal automation while fiddling around with the O365 installation on my personal Linux machine.
At a later point I got the idea to share my work by publishing my personal scripts after making them a tiny bit more user-friedly.

Some special thanks to ruados, whos [article](https://ruados.github.io/articles/2021-05/office365-wine) especially helped me creating these scripts.
## Known flaws
Before use, be aware of some known flaws:
* The scripts do have almost no error handling, neither for script or machine errors nor for input errors.
* They're currently only made for Debian-based Distributions, mainly because of included apt-get commands and apt source installation.
* The scripts include commands with some hardcoded 'Office16' clauses, which might get outdated, when Microsofts releases a new Office365 release. (...and many other reasons, why the whole scrips basically could get outdated quickly, of course).
* The scripts were manually tested and only with my private desktop machine, which uses [siduction](https://siduction.org/), a Debian Sid based distribution (state 2022-04-25 a Debian "bookworm" release). So, you might encounter unknown issues with other distributions or even Debian release versions.
* The scripts were not designed for Linux starters, though they still might help them. I'd still like to advise: Do not use one of the scripts, if you do not understand, what the scripts are doing in and with your system.
* The scripts currently install only the 32 Bit version of O365. For people, who need the 64 Bit Office, the scripts are currently not suited.
* I'm not a real programmer myself, so be cautious, that the script might not follow all otherwise vastly established programming principles.
* Currently, even if the script and installation runs fine, you probably cannot log into your Microsoft account with Office (outdated browser warning).
* ...and maybe many more...
## Variants
There're two variants of the installation script:
* one utilizes the [Office Deployment Tool (ODT)](https://support.microsoft.com/en-us/office/use-the-office-offline-installer-f0a85fe7-118f-41cb-a791-d59cef96ad1c#OfficePlans=signinorgid),
* the other one uses the 32 bit Retail installer file, which you can download via your [private Microsoft account](https://account.microsoft.com/services/microsoft365/details) or [business Microsoft account](https://portal.office.com/account/?ref=MeControl#installs)
  * On both it is required, that the browser agent is showing "Windows" as OS to the opened websites, otherwise the download won't be offered.
## Additional hints
* Neither the ODT nor the Retail installer exe-file will be downloaded by the script. You have to download them beforehand.
* I'd recommend using the Retail installer script, because O365 generally is pre-activated. The O365 installed via ODT probably cannot be activated, due to the missing login functionality already mentioned above.
* If something goes wrong during script runtime, you might just abort and restart it. The formerly created wine prefix with the same name will be deleted by the script.
* I also added a toggle to choose between installing winhttp over wineprefix or not, which seems to be necessary in some constellations, but also seems to break the installation in others. The default though should be `winhttp=no`.
## Usage
Download the script you want to use, or checkout the whole git repo. The usage of the scripts is pretty similar:
```
# Make the script executable once
chmod +x script-name.sh
# Start the script
./script-name.sh distribution distrireleasename prefixname winhttp=[yes|no] pathtoexefile
```

Or more specifically...

For a Debian bookworm-based system, the ODT script could be started e. g. like this:
```
./wine-prefix_O365_32bit_odt.sh debian bookworm msoffice365 winhttp=no ~/Downloads/officedeploymenttool_15028-20160.exe
```
For a Debian bookworm-based system the Retail script then could be started e. g. like this:
```
./wine-prefix_O365_32bit_retail.sh debian bookworm msoffice365 winhttp=no ~/Downloads/OfficeSetup32.exe 
```
For an Ubuntu 22.04 LTS-based system the Retail script then could be started e. g. like this (not tested):
```
./wine-prefix_O365_32bit_retail.sh ubuntu jammy msoffice365 winhttp=no ~/Downloads/OfficeSetup32.exe
```
...and so on and so forth...
