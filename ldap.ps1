<#
.SYNOPSIS
Позволяет получать название компьютеров по ФИО или username пользователей.
.Description
Позволяет получать название компьютеров по ФИО или username пользователей.
.PARAMETER ping
Необязательная опция, если добавлена то проводится проверка доступности хостов (увеличевает время выполнения), иначе отключает проверку доступности.
.PARAMETER data
Путь к файлу со списком ФИО.
.PARAMETER group
Группировака по пользователю (все компьютеры в одном поле)
.PARAMETER groups
Необязательный параметр, разделитель для групп (по умпочанию Delimeter=$null - не выводит группы).
.PARAMETER csv 
Необязательный параметр, путь к файлу для вывода в формате CSV.
.PARAMETER delimeter
Необязательный параметр, разделитель для формата CSV (по умпочанию Delimeter=",").
.PARAMETER ADCModule
Путь к модулю если он не установлен.
.LINK
https://github.com/samratashok/ADModule
.EXAMPLE
.\ldap.ps1 -data users.txt
.EXAMPLE
.\ldap.ps1 -data "C:\Users\User1\Documents\ФИО.txt" -csv "mylist.csv" -ping -group
.EXAMPLE
.\ldap.ps1 -data "C:\Users\User1\Documents\ФИО.txt" -csv "mylist.csv" -delimeter ";"
.EXAMPLE
.\ldap.ps1 -data "C:\Users\User1\Documents\ФИО.txt" -csv "mylist.csv" -delimeter ";" -ADModule "Y:\Corps\СИБ\_ДИБ\Личные папки\Байрашный\powershell\ADModule\Microsoft.ActiveDirectory.Management.dll"
#>

param(
    [Parameter(Mandatory=$true)][String]$data,
    [string]$csv,
    [string]$delimeter=';',
    [switch]$ping=$false,
    [switch]$group=$false,
	[String]$groups=$null,
    [string]$ADCModule="Y:\Corps\СИБ\_ДИБ\Личные папки\Байрашный\powershell\ADModule\Microsoft.ActiveDirectory.Management.dll"
)

function Fix-Name($name) {
    return $(if($name -is [String]) { $name.trim() -replace "  "," " } else { $null })
}

function isAvailable($computer){
	return Test-Connection -ComputerName $computer -Quiet -Count 1 -ErrorAction SilentlyContinue
}

filter Ping {
	(($_.Name) -and (isAvailable -computer $_.Name)) -or (($_.IPAddress) -and (isAvailable -computer $_.IPAddress))
}

function Get-ADComputers($ping=$true,$data,$group=$false,$csv,$delimeter=';',$groups=$null){
    if (Get-Module -All -Name *ActiveDirectory*) {
		$list=foreach($fio in Get-Content $data -Encoding UTF8 | Where-Object {$_ -match '\w{4}.*'}) {
			$users=$null
            if($fio -match '^[a-zA-Z]+$'){
                $users=Get-ADUser $fio -Properties SamAccountName,Enabled,EmailAddress,CanonicalName,UserPrincipalName,Name,Title,MemberOf
            }else{
                $users=Get-ADUser -LDAPFilter "(|(Name=*$fio*)(dispalyName=*$fio*)(cn=*$fio*)(UserPrincipalName=*$fio*))" -SearchBase "DC=oek,DC=ru" -Properties SamAccountName,Enabled,EmailAddress,CanonicalName,UserPrincipalName,Name,Title,MemberOf
            }
            if($users){
				foreach($user in $users){
					$s=if($user){($user.CanonicalName -replace "/$($user.Name)","").split("/")}else{@()}
					$account=$user.SamAccountName
					$computers=Get-ADComputer -LDAPFilter "(|(DNSHostName=*$account*)(Name=*$account*))" -SearchBase "DC=oek,DC=ru" -Properties *
					if($null -eq $computers){
						[PSCustomObject]@{
							account = $user.SamAccountName
							enabled = $user.Enabled
							email = $user.EmailAddress
							name = Fix-Name $user.Name
							post = Fix-Name $user.Title
							office = Fix-Name $s[2]
							departament = Fix-Name $s[3]
							division = Fix-Name $s[4]
							groups = if($groups){($user.MemberOf | ForEach-Object { ($_ -replace "(CN|OU|DC)=","") -replace ",","/" }) -join $groups}else{$null}
						}
					}elseif ($group) {
						[PSCustomObject]@{
							account = $user.SamAccountName
							enabled = $user.Enabled
							email = $user.EmailAddress
							name = Fix-Name $user.Name
							post = Fix-Name $user.Title
							office = Fix-Name $s[2]
							departament = Fix-Name $s[3]
							division = Fix-Name $s[4]
							groups = if($groups){($user.MemberOf | ForEach-Object { ($_ -replace "(CN|OU|DC)=","") -replace ",","/" }) -join $groups}else{$null}
							available = if($ping){($computers | Ping | foreach-object  {$_.DNSHostName}) -join ", "}else{"Not checked"}
							hosts = ($computers | foreach-object  {$_.DNSHostName}) -join ", "
							ips = if($ping){($computers | foreach-object  {(Resolve-DNSName $_.DNSHostName -ErrorAction SilentlyContinue).IPAddress}) -join ", "}else{"Not checked"}
						}
					}else{
						foreach($computer in $computers){
							[PSCustomObject]@{
								account = $user.SamAccountName
								enabled = $user.Enabled
								email = $user.EmailAddress
								name = Fix-Name $user.Name
								post = Fix-Name $user.Title
								office = Fix-Name $s[2]
								departament = Fix-Name $s[3]
								division = Fix-Name $s[4]
								groups = if($groups){($user.MemberOf | ForEach-Object { ($_ -replace "(CN|OU|DC)=","") -replace ",","/" }) -join $groups}else{$null}
								available = if($ping){isAvailable -computer $computer.DNSHostName}else{"Not checked"}
								host = $computer.DNSHostName
								ip = if($ping){(Resolve-DNSName $computer.DNSHostName -ErrorAction SilentlyContinue).IPAddress}else{"Not checked"}
							}
						}
					}
				}
			}else{
				[PSCustomObject]@{
					name = $fio
				}
			}
		}
		
		if($group){
			$list=$list | Select-Object -Property account,enabled,email,name,post,office,departament,division,groups,available,hosts,ips
		}else{
			$list=$list | Select-Object -Property account,enabled,email,name,post,office,departament,division,groups,available,host,ip
		}
		if($csv -gt 0){
			$list | Export-CSV $csv -NoTypeInformation -Delimiter $delimeter -Encoding utf8
		}else{
			$list | Format-Table
		}
    }else{
        Write-Output 'Module "Microsoft.ActiveDirectory.Management" does not imported.'
        Write-Output 'Download this module from: https://github.com/samratashok/ADModule'
        Write-Output 'And run with -ADModule "C:\...\Microsoft.ActiveDirectory.Management.dll"'
    }
}


if (!(Get-Module -All -Name *ActiveDirectory*)) {
    Import-Module $ADCModule 
}

Get-ADComputers -ping $ping -data $data -group $group -csv $csv -delimeter $delimeter -groups $groups
