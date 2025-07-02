<#
help/man command
	-to help undestand/search for PS commands
	-use wild card statements (*) in case you have keywords that you want to search for (eg. *Event*)
	-the brackets [] when looking at the help syntax are meant to indicate that the particular parameter
	or name that needs to be included is optional (eg. [[-Path] <string[]>]
	-type help [commandName] -full to get the description as well as a further breakdown of the parameters
#>

Get-PSDrive #to get info about the available drives in current PS
man dir -full >> dir-manual.txt #man (seen in unix mainly) and help are both aliases
help *reboot* #searches for commands that have the word reboot or have a desc that has the keyword in it
