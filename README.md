# OOOBackupRestore
Out Of Office Backup and Restore for Exchange 2010

This project was born after I was asked to replace all users' OOO (Out Of Office) messages with a unified Christmas message, preserving specific custom user settings and restoring those individual specific out of office settings afterwards.

I searched online for tools and options, then I ended up writing my own Powershell scripts, to accomplish the above.

The project is fairly simple, and it is comprised of three scripts; their scope of users is defined by the username list file, commonly used by all the .ps1 scripts.

1) Backup script
This script carries out a complete backup of all selected users out of office settings, into a CSV file.

2) Standard OOO Message setup script
This script overwrites all selected users mailboxes with a new, custom message (It can be HTML as well)

3) Restore script
This script restores the OOO setup previously saved into a CSV file, overwriting the current mailbox OOO Settings.

This has been written and tested on Microsoft Exchange 2010, but it should be good to go for Exchange 2013 and 2016.
