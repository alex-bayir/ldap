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
.\ldap.ps1 -data "C:\Users\User1\Documents\ФИО.txt" -csv "mylist.csv" -delimeter ";" -ADModule "C:\Users\User1\Downloads\ADModule\Microsoft.ActiveDirectory.Management.dll"
#>

param(
    [Parameter(Mandatory=$true)][String]$data,
    [string]$csv,
    [string]$delimeter=$null,
    [switch]$ping=$false,
    [switch]$group=$false,
    [string]$ADCModule="Y:\Corps\СИБ\Байрашный\ADModule\Microsoft.ActiveDirectory.Management.dll"
)

function isAvailable($computer){
	return Test-Connection -ComputerName $computer -Quiet -Count 1 -ErrorAction SilentlyContinue
}

filter Ping {
	(($_.Name) -and (isAvailable -computer $_.Name)) -or (($_.IPAddress) -and (isAvailable -computer $_.IPAddress))
}

function Get-ADComputers($ping=$true,$data,$group=$false,$csv,$delimeter=$null){
    if (Get-Module -All -Name *ActiveDirectory*) {
		$list=foreach($fio in Get-Content $data -Encoding UTF8 | Where-Object {$_ -match '\w{4}.*'}) {
			$users=$null
            if($fio -match '^[a-zA-Z]+$'){
                $users=Get-ADUser $fio
            }else{
                $users=Get-ADUser -LDAPFilter "(|(Name=*$fio*)(dispalyName=*$fio*)(cn=*$fio*)(UserPrincipalName=*$fio*))" -SearchBase "DC=oek,DC=ru" -Properties *
            }
            if($users){
				foreach($user in $users){
                    #Write-Output $user
			        $account=$user.SamAccountName
			        $computers=Get-ADComputer -LDAPFilter "(|(uid=*$account*)(dispalyName=*$account*)(cn=*$account*)(sn=*$account*))" -SearchBase "DC=oek,DC=ru" -Properties *  
                    if($group){
						[PSCustomObject]@{
							Name           = $user.Name
							AccountName    = $user.SamAccountName 
							Available      = if($ping){($computers | Ping | foreach-object  {$_.Name}) -join ", "}else{"Not checked"}
							Hosts          = ($computers | foreach-object  {$_.Name}) -join ", "
                            IPs            = if($ping){($computers | foreach-object  {(Resolve-DNSName $_.Name -ErrorAction SilentlyContinue).IPAddress}) -join ", "}else{"Not checked"}
							UserName       = $user.UserPrincipalName -replace "@oek.ru", ""
							UserEnabled    = $user.Enabled
							CN             = $user.DistinguishedName
						}
                    }else{
                        if($computers){
							foreach($computer in $computers){
								[PSCustomObject]@{
									Name           = $user.Name
									AccountName    = $user.SamAccountName 
									Available      = if($ping){isAvailable -computer $computer.Name}else{"Not checked"}
									Host           = $computer.Name
									IP             = if($ping){(Resolve-DNSName $computer.Name -ErrorAction SilentlyContinue).IPAddress}else{"Not checked"}
									UserName       = $user.UserPrincipalName -replace "@oek.ru", ""
									UserEnabled    = $user.Enabled
									CN             = $user.DistinguishedName
								}
							}
                        }else{
							[PSCustomObject]@{
								Name           = $user.Name
								AccountName    = $user.SamAccountName
                                UserName       = $user.UserPrincipalName -replace "@oek.ru", ""
								UserEnabled    = $user.Enabled
								CN             = $user.DistinguishedName
                            }
						}
					}
                }
            }else{
                [PSCustomObject]@{
				    Name           = $fio
                }
            }
		}
		
		$delimeter=$(if ($delimeter.Length -gt 0) {$delimeter} else {","})
		$list=$list | Select-Object -Property Name,AccountName,Available,Hosts,IPs,UserName,UserEnabled,CN
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

Get-ADComputers -ping $ping -data $data -group $group -csv $csv -delimeter $delimeter
