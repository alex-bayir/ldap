del out.csv
powershell -noexit ".\ldap.ps1 -data users.txt -group -csv out.csv -delimeter ';'"