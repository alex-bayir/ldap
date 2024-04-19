param(
    [Parameter(Mandatory=$true)][String]$data,
    [string]$server=$null,
    [string]$csv,
    [string]$delimeter=$null,
    [string]$ADCModule="Y:\Corps\СИБ\_ДИБ\Личные папки\Байрашный\powershell\ADModule\Microsoft.ActiveDirectory.Management.dll"
)

if (!(Get-Module -All -Name *ActiveDirectory*)) {
    Import-Module $ADCModule
}
function Write-Error($message) {
    [Console]::ForegroundColor = 'red'
    [Console]::Error.WriteLine($message)
    [Console]::ResetColor()
}
function Get-Posts ([Parameter(Mandatory=$true)][String]$fio,[Object[]]$users=@()){
    $response = Invoke-WebRequest -Method Post 'https://uneco2.ru/local/ajax/header_search.php' -Body @{ 'phrase' = $fio } -ContentType 'application/x-www-form-urlencoded; charset=UTF-8'
    $employees = $response.ParsedHtml.getElementsByClassName("popup__employees-list")
    if($employees.Length -eq 0){
        Write-Error "Не найден: ""$fio"""
        $users | ForEach-Object {
            [PSCustomObject]@{
                url = $null
                account = $_.SamAccountName
                enabled = $_.Enabled
                name = $fio
                post = $null
                division = $null
                1 = $null
                2 = $null
                3 = $null
                4 = $null
            }
        }
    }else{
        $employees[0].getElementsByTagName("a") | ForEach-Object {
            [PSCustomObject]@{
                name = $_.textContent -replace " Написать сообщение",""
                url = "https://uneco2.ru/"+$_.pathname
            }
        } | Where-Object { $_.name -match $fio } | ForEach-Object {
            $url=$_.url
            $html=(Invoke-WebRequest $url).ParsedHtml
            $name = $html.getElementsByClassName("content-item__title")[0].textContent.trim() -replace "[^а-яёА-ЯЁ a-zA-Z]",""
            $post = $html.getElementsByClassName("content-item__text")[0].textContent.trim()
            $arr = $html.getElementsByClassName("arrow-left") | Where-Object { $_.children.Length -gt 0 -and $_.children } | ForEach-Object { $_.children[0].title } | Where-Object { $_ -ne $null -and $_ -ne "" -and $_ -ne "Структура компании" } | ForEach-Object { $_.trim() }
            if(($arr -is [String[]]) -eq $false){ $arr=@($arr) }
            $users | Where-Object { ($_.name -replace "ё","е") -eq ($name -replace "ё","е") } | ForEach-Object {
                [PSCustomObject]@{
                    url = $url
                    account = $_.SamAccountName
                    enabled = $_.Enabled
                    name = $name
                    post = $post
                    division = $arr[-1]
                    1 = $arr[0]
                    2 = $arr[1]
                    3 = $arr[2]
                    4 = $arr[3]
                }
            }
        }
    }
}

$content = Get-Content $data -Encoding UTF8
$i=0; $length = $content.Length
$users = $content | ForEach-Object {
    $i+=1; Write-Host -NoNewLine "" "`r$i/$length "
    if($_ -match '^[a-zA-Z]+$'){
        $user = Get-ADUser $_
        Get-Posts -fio $user.name -users @($user)
    } else {
        if($server){
            $users = Get-ADUser -Server $server -LDAPFilter "(|(Name=*$_*)(dispalyName=*$_*)(cn=*$_*)(UserPrincipalName=*$_*))" -SearchBase "DC=oek,DC=ru" -Properties *
        }else{
            $users = Get-ADUser -LDAPFilter "(|(Name=*$_*)(dispalyName=*$_*)(cn=*$_*)(UserPrincipalName=*$_*))" -SearchBase "DC=oek,DC=ru" -Properties *
        }
        Get-Posts -fio $_ -users $users
    }
}
Write-Host ""
if($csv -gt 0){
    $users | Export-CSV $csv -NoTypeInformation -Delimiter $delimeter -Encoding utf8
}else{
    $users | Format-Table
}