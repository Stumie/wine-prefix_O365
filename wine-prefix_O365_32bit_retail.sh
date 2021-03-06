#!/bin/bash

if [ "$(id -u)" = "0" ]; then
  echo "This script should not be run as root." 1>&2
  exit 1
fi

if (( $# != 7 )); then
  echo "Not all or too many parameters provided"
  exit 1
fi

# Set variables
DISTRIRELEASE=$1
DISTRIVERSION=$2
WINEBRANCHNAME=$3
WINEPREFIXNAME=$4
WINENATIVENESS=$5
WINHTTPYESNO=$6
WINESWSETUPPATH=$7

WINEPREFIXPATH=~/.wine-prefix/$WINEPREFIXNAME
WINEARCH=win32

# Download and install requirements via APT
sudo dpkg --add-architecture i386
sudo apt-get update
sudo apt-get upgrade -y
sudo apt-get install sudo vim gcc make perl wget dos2unix software-properties-common -y
cd ~
wget -nc https://dl.winehq.org/wine-builds/winehq.key
sudo mv winehq.key /usr/share/keyrings/winehq-archive.key
wget -nc https://dl.winehq.org/wine-builds/$DISTRIRELEASE/dists/$DISTRIVERSION/winehq-$DISTRIVERSION.sources
sudo mv winehq-$DISTRIVERSION.sources /etc/apt/sources.list.d/
sudo apt-get update

if [ $WINEBRANCHNAME = 'stable' ] || [ $WINEBRANCHNAME = 'staging' ] || [ $WINEBRANCHNAME = 'devel' ]; then
  sudo apt-get install --install-recommends winehq-$WINEBRANCHNAME -y
else 
  echo "ABORT: Provide valid winebranch statement within script parameters"
  exit 1
fi

sudo apt-get install winetricks cabextract -y
sudo apt-get install p11-kit p11-kit-modules winbind samba smbclient -y

# Kill and remove existing prefix with the same name
WINEPREFIX=$WINEPREFIXPATH WINEARCH=$WINEARCH wineboot --force --kill
WINEPREFIX=$WINEPREFIXPATH WINEARCH=$WINEARCH winetricks -q annihilate
rm -rf $WINEPREFIXPATH

# Create the wine prefix and do a "first boot"
mkdir -p $WINEPREFIXPATH
WINEPREFIX=$WINEPREFIXPATH WINEARCH=$WINEARCH wineboot -i

# Install Dependencies within Wine prefix via Winetricks
case $WINENATIVENESS in
  'level=1')
    # Minimal set of winetricks
    WINEPREFIX=$WINEPREFIXPATH WINEARCH=$WINEARCH winetricks -q win7 msxml3 msxml6 riched20 riched30
    ;;
  'level=2')
    # Compatibility set of winetricks
    WINEPREFIX=$WINEPREFIXPATH WINEARCH=$WINEARCH winetricks -q win7 corefonts urlmon msxml3 msxml6 riched20 riched30 gdiplus
    ;;
  'level=3')
    # More native set of winetricks
    WINEPREFIX=$WINEPREFIXPATH WINEARCH=$WINEARCH winetricks -q win7 corefonts urlmon msxml3 msxml6 riched20 riched30 gdiplus msftedit dotnet20 dotnet472 vcrun2019
    ;;
  *)
    echo "ABORT: Provide valid nativeness level statement within script parameters"
    exit 1
    ;;
esac

# Workaround: While winhttp seems to be necessary in some constellations, it seems to also break the installation in others.
# This offers an option to choose within the parameters, so without editing the script itself.
if [ $WINHTTPYESNO = 'winhttp=yes' ]; then
  WINEPREFIX=$WINEPREFIXPATH WINEARCH=$WINEARCH winetricks -q winhttp
elif [ $WINHTTPYESNO = 'winhttp=no' ]; then
  break
else 
  echo "ABORT: Provide valid winhttp statement within script parameters"
  exit 1
fi

# Workaround: Add Wine registry keys for graphics fixes and DLL overrides
WINEPREFIX=$WINEPREFIXPATH WINEARCH=$WINEARCH wine regedit /E ~/export.reg "HKEY_CURRENT_USER\Software\Wine\Debug"
dos2unix -o ~/export.reg
cat ~/export.reg | head -2 > ~/import.reg
cat << EOF >> ~/import.reg
[HKEY_CURRENT_USER\Software\Wine\Direct2D]
"max_version_factory"=dword:0
[HKEY_CURRENT_USER\Software\Wine\Direct3D]
"MaxVersionGL"=dword:00030002
[HKEY_CURRENT_USER\Software\Wine\DllOverrides]
"sppc"=""
EOF
unix2dos -o ~/import.reg
WINEPREFIX=$WINEPREFIXPATH WINEARCH=$WINEARCH wine regedit ~/import.reg
rm ~/export.reg
rm ~/import.reg

# Reboot Wine prefix
WINEPREFIX=$WINEPREFIXPATH WINEARCH=$WINEARCH wineboot -u
WINEPREFIX=$WINEPREFIXPATH WINEARCH=$WINEARCH wineboot -r

# Make sure Wine prefix is set to Windows 10 before first setup
WINEPREFIX=$WINEPREFIXPATH WINEARCH=$WINEARCH winecfg -v win10

printf '%s\n' "------------------------------------------------------------"
printf '%s\n' "The script will now start the Office installer. If any wine error window occurs, just click it away."
printf '%s\n' "The first Office installation will probably break around 97%. Wait for the crash, close the crash and setup window and press here any key to continue: A second installer will start."
printf '%s\n' "------------------------------------------------------------"

# Start Office setup for the first run
WINEPREFIX=$WINEPREFIXPATH WINEARCH=$WINEARCH wine $WINESWSETUPPATH

# Let script wait until user interacts (until Office installation finished.)
while [ true ] ; do
  read -t 3 -n 1
  if [ $? = 0 ] ; then
    break
  else
    printf '%s\n' "------------------------------------------------------------"
    printf '%s\n' "If any wine error window occurs, just click it away."
    printf '%s\n' "The first Office installation will probably break around 97%. Wait for the crash, close the crash and setup window and press here any key to continue: A second installer will start."
    printf '%s\n' "------------------------------------------------------------"
  fi
done

# Make sure Wine prefix is set back to Windows 7 after setup
WINEPREFIX=$WINEPREFIXPATH WINEARCH=$WINEARCH winecfg -v win7

# Start Office setup for the second run
WINEPREFIX=$WINEPREFIXPATH WINEARCH=$WINEARCH wine $WINESWSETUPPATH

# Let script wait until user interacts (until Office installation finished.)
while [ true ] ; do
  read -t 3 -n 1
  if [ $? = 0 ] ; then
    break
  else
    printf '%s\n' "------------------------------------------------------------"
    printf '%s\n' "If any wine error window occurs, just click it away."
    printf '%s\n' "The installation takes longer again at around 97%, but should't break."
    printf '%s\n' "When the Office installation finished (Pause symbol visible), close the setup window and press here any key to continue."
    printf '%s\n' "------------------------------------------------------------"
  fi
done

# Workaround: Symlink creation seems to be broken during installation, which is the reason, why DLLs have to be copied manually.
cp -fv $WINEPREFIXPATH/drive_c/Program\ Files/Common\ Files/Microsoft\ Shared/ClickToRun/*.dll $WINEPREFIXPATH/drive_c/Program\ Files/Microsoft\ Office/root/Office1*/
cp -fv $WINEPREFIXPATH/drive_c/Program\ Files/Common\ Files/Microsoft\ Shared/ClickToRun/*.dll $WINEPREFIXPATH/drive_c/Program\ Files/Microsoft\ Office/root/Client

printf '%s\n' "------------------------------------------------------------"
printf '%s\n' "Script finished."
printf '%s\n' "Start Office applications e. g. like this:"
printf '%s\n' "WINEPREFIX=$WINEPREFIXPATH WINEARCH=$WINEARCH wine $WINEPREFIXPATH/drive_c/Program\ Files/Microsoft\ Office/root/Office16/EXCEL.EXE"
printf '%s\n' "------------------------------------------------------------"
