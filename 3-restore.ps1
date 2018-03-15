# This restores the autoreply settings from a backup in saj.csv
#
# Fetch the spreadsheet
$bigtable=ipcsv c:\pst\backupfile.csv

# Warn about overwriting
write-warning "This will overwrite the autoreply settings for the users listed in backupfile.csv " -warningaction Inquire

# Restore each mailbox depending on settings
foreach ($item in $bigtable)
{
    # Print which user is being processed
    write-host "Processing: " $($item.identity)
    
    # If the mailbox is scheduled then restore all values available
    if ($($item.autoreplystate) -eq "Scheduled") 
    {
        # Convert dates/times to correct format
        $st = [DateTime]::Parse($item.starttime)
        $et = [DateTime]::Parse($item.endtime)
        # Restore all values
        set-mailboxautoreplyconfiguration -identity $($item.identity) -starttime "$st" -endtime "$et" -internalmessage $($item.internalmessage) -externalmessage $($item.externalmessage) -autoreplystate $($item.autoreplystate)
    }
    else
    {
        # Restore Messages and state, dates are irrelevant so set static ones
        set-mailboxautoreplyconfiguration -identity $($item.identity) -starttime "01/02/2016 1:00 PM" -endtime "01/02/2016 2:00 PM" -internalmessage $($item.internalmessage) -externalmessage $($item.externalmessage) -autoreplystate $($item.autoreplystate)
    }
}
    