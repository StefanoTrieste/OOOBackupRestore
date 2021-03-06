# Backup of system autoreply settings
# List of identities involved must be available in identities.csv, either username or name surname are ok.
#
# SL 8/12/2015

write-host "Processing identities.csv ..." -nonewline
ipcsv c:\pst\identities.csv | get-mailboxautoreplyconfiguration | select identity, starttime, endtime, autoreplystate, internalmessage, externalmessage| epcsv c:\pst\backupfile.csv
write-host "Done"
