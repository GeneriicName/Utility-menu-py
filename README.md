# Utility-menu-py
A utility menu i made while working as an help desk, was made in between calls, so i'm aware its far from perfect.

***you will need to configure the config file before starting it up***, i added some comments as for how to configure it and with some examples.

the features are as for the following:
***Clean space from remote computers***, you can configure which directories to clean up via the config file, as well as if to delete the search edb file of windows
***Get network printers from the computer***, including printers installed via print servers, TCP/IP and WSD (will retrive the IP, and tell you if its found on any of the print servers)
***Delete the OST file from the remote computer*** 
***Reset print spooler*** 
***Fix internet explorer***, depending on your OS version this might not work
***Fix cockpit printers***, via deleting the appropriate registry keys - usefull only if your org uses jetro cockpit
***Fix 3 languages bug***, fixes a bug when the same language get displayed twice
***Delete users folders***, to clean up space - if the user is in your domain it'll show its display name - so it's clearer which users are you deleting

***Display a bunch of information on the computer and user***
1. The script will show the user status (active, locked, disabled, expired, or password expired)
2. Same goes for the computer, it will show space in C or D disk, uptime, current user, if the computer is online/offline

**You will need to have a user.txt file containing the computername from which the last user has logged on to if you want to be able to use this script with username as well as hostnames,
this could be easily achived by a simple batch logon script/GPO/task, and the location to which the files are dumped need to be configured in the config.json file**

**I'm fully aware this is'nt a completed program, i uploaded it with some features missing (only relevent for the company i'm working for, and i wouldn't want to release sensetive information
such as AD groups, users, paths etc) but this script contains a few nice features such as retriving network printers from a remote computer, displaying user and computer information, multithreaded file deletion,
and a fairly simple GUI, all of which could be taken and modfied for your own project**

I wont be updating this script any further, but if you need help to understand why i did somthing in my script or how it works, feel free to ask.
