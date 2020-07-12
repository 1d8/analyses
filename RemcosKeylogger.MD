# Remcos RAT Keylogger

## Overview
* [sample here](https://bazaar.abuse.ch/sample/6ad868658b3f50bfde225d52382d33a6027e8344592fb6fa296b8516d9d00f4c/) 

* Password to zip is **infected**

* First seen: June 10, 2020

* Sha256 hash: 6ad868658b3f50bfde225d52382d33a6027e8344592fb6fa296b8516d9d00f4c

* filetype: exe


# Analysis

This keylogger doesn't require any special action to trigger it to run, it constantly runs in the background & logs all events created by the user.

After running the .exe file, we see that it creates a directory in the %USERPROFILE%\AppData\Roaming directory named **remcos**, it is here that the logs will be stored *Ignore the name collision result, it only occurred because I ran the malware twice*:
 
![](/pic-set2/img1.png)

A new file is also created in the remcos directory, it's named **logs.dat** :

![](/pic-set2/img2.png)

![](/pic-set2/img3.png)

The keylogger itself is quite noisy, it edits the **logs.dat** file every time a new event occurs on the system (switching tasks, typing something in the keyboard, launching a new process, etc) which is likely to be less than every second if the computer is in use:

![](/pic-set2/img4.png)

One thing I found interesting was that the malware appears to interact with `HKCU\Software\Buddy-6SE1KQ\` before editing the log file:

![](/pic-set2/img5.png)

Here's an example of the activity that is logged:

![](/pic-set2/img6.png)
