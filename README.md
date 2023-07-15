# Utility-menu.py
A utility menu I made while working as a help desk. It was made in between calls, so I'm aware it's far from perfect.

***You will need to configure the config file before starting it up.*** I added some comments on how to configure it and provided some examples.

The features are as follows:

***Clean space from remote computers:*** You can configure which directories to clean up via the config file, as well as whether to delete the Windows search edb file.

***Get network printers from the computer:*** This includes printers installed via print servers, TCP/IP, and WSD. It will retrieve the IP and tell you if it's found on any of the print servers.

***Delete the OST file from the remote computer.***

***Reset print spooler.***

***Fix Internet Explorer:*** Depending on your OS version, this might not work.

***Fix cockpit printers:*** This deletes the appropriate registry keys. Useful only if your organization uses Jetro Cockpit.

***Fix 3 languages bug:*** Fixes a bug when the same language is displayed twice.

***Delete users folders:***, Choose users to delete their folders in order to cleans up space. If the user is in your domain, it'll show their display name to make it clearer which users you are deleting, will exclude the current user of the remote computer, as well as additonal users that you can configure in the config.json file

***The script also displays a bunch of information on the computer and user:***
1. The script shows the user status (active, locked, disabled, expired, or password expired).
2. Similarly, it shows the computer status, including space in the C or D disk, uptime, current user, and whether the computer is online/offline.
   
**To use this script with usernames as well as hostnames, you will need a user.txt file containing the computer name from which the last user has logged on. You can easily achieve this with a simple batch logon script/GPO/task, and the location to which the files are dumped needs to be configured in the config.json file.**

**I'm fully aware that this isn't a completed program. I uploaded it with some features missing (only relevant for the company I'm working for), and I wouldn't want to release sensitive information such as AD groups, users, paths, etc. However, this script contains a few nice features such as retrieving network printers from a remote computer, displaying user and computer information, multithreaded file deletion, and a fairly simple GUI. You can take and modify these features for your own project.**

***Additional info: This script features support for logging, although not fully. However, points of the script that are most likely prone to exceptions are covered. Also, the "assets" folder contains all the images for the GUI application. You can paste a username/computer name/IP address or network path of a printer in the entry box. I didn't split the script into multiple files as I compile it to an executable using Nuitka for ease of distribution to other team members.***

I won't be updating this script any further, but if you need help understanding why I did something in my script or how it works, feel free to ask.
![image](https://github.com/GeneriicName/Utility-menu-py/assets/139624416/e8cf7404-8e4d-41a6-ae73-cb231ebf6c0b)

