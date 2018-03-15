#Replaces users' settings with the festive message in xmasmessage.txt
#
#
# Fetch list of mailboxes involved, from the backup spreadsheet. This needs to be current.
# Have you just run export.ps1 ? Make sure you have a backup!
$usr=ipcsv c:\pst\backupfile.csv

# Retrieve the common message
$msg=get-content C:\PST\custommessage.txt

# Set up start and end dates/times with correct format, make sure your PC is set to the correct time zone. The example below is valid for Australia/English.
$st = [DateTime]::Parse("22/12/2017 4:00 PM")
$et = [DateTime]::Parse("8/1/2018 8:00 AM")

# Warn user about possible loss of data
write-warning "This will overwrite the autoreply settings for the users listed in identities.csv make sure you have a backup before continuing!" -warningaction Inquire

# Process the list and set out of office
foreach ($item in $usr)
{
    write-host "Processing: " $($item.identity)
    set-mailboxautoreplyconfiguration -identity $($item.identity) -starttime "$st" -endtime "$et" -internalmessage $msg -externalmessage $msg -autoreplystate "Scheduled"
}
    