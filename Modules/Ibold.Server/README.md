# Pre-Requisites
1.  Powershell v5+.
2.  ActiveDriectory PowerShell Module.
3.  This script must be run on an on-premise server and should be run on a DC.
4.  This script does not require an internet connection.

# Running the script
1.  Save the script to your computer and run the script via CLI
2.  There are no params. Just fire up the PS1

# What it does
1.  Gathers the following information:
    - Domain and Forest architecture
    - Domain trusts, sites, controllers and FSMO role holders
    - OU Structure
    - AD Objects (users, computers, groups, privileged groups, domain admins, service acccounts (user and gMSA))
2.  Fully backs up GPOs, structure and objects (scripts, batch files, etc.) and generates an HTML report.
3.  Password security objects
4.  Optionally installed AD features
5.  DNS, A-Records, server zones, and scavenging
6.  Outputs all of the information to individual CSVs, and for information that doesn't make sense to have its own CSV, it is appended to a text file with readable
    formatting.
7.  Compressed the GPO backup folder into a ZIP file and deletes the uncompressed folder.

# What it does *not* do

# To-Do:
