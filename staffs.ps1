param(
    [Parameter(Mandatory=$true)][String]$data,
    [string]$csv,
    [string]$delimeter=$null,
    [string]$ADCModule="Y:\Corps\СИБ\Байрашный\ADModule\Microsoft.ActiveDirectory.Management.dll"
)

if (!(Get-Module -All -Name *ActiveDirectory*)) {
    Import-Module $ADCModule 
}
function Write-Error($message) {
    [Console]::ForegroundColor = 'red'
    [Console]::Error.WriteLine($message)
    [Console]::ResetColor()
}
function Get-Posts {
    param (
        [Parameter(Mandatory=$true)][String]$fio,
        $users=@()
    )
    $response = Invoke-WebRequest -Method Post 'http://uneco2.ru/local/ajax/header_search.php' -Body @{ 'phrase' = $fio } -ContentType 'application/x-www-form-urlencoded; charset=UTF-8'
    $employees = $response.ParsedHtml.getElementsByClassName("popup__employees-list")
    if($employees.Length -eq 0){
        Write-Error "Не найден: ""$fio"""
        $user = if($users){ if($users -is [string[]]){ $users[0] }else{ $users } }else{ $null }
        [PSCustomObject]@{
            account = $user.SamAccountName
            enabled = $user.Enabled
            name = $fio
        }
    }else{
        $employees[0].getElementsByTagName("a") | ForEach-Object {
            $url = "https://uneco2.ru/"+$_.pathname
            $html=(Invoke-WebRequest $url).ParsedHtml
            $name = $html.getElementsByClassName("content-item__title")[0].textContent.trim() -replace "[^а-яёА-ЯЁ a-zA-Z]",""
            $post = $html.getElementsByClassName("content-item__text")[0].textContent.trim()
            $arr=$html.getElementsByClassName("arrow-left") | Where-Object { $_.children.Length -gt 0 -and $_.children } | ForEach-Object { $_.children[0].title } | Where-Object { $_ -ne $null -and $_ -ne "" -and $_ -ne "Структура компании" }
            if(($arr -is [string[]]) -eq $false){ $arr=@($arr) }
            $user = $users | Where-Object { $_.name -eq $name }
            $user = if($user){ $user }else{ $null }
            [PSCustomObject]@{
                url = $url
                account = $user.SamAccountName
                enabled = $user.Enabled
                name = $name
                post = $post
                1 = $arr[0]
                2 = $arr[1]
                3 = $arr[2]
                4 = $arr[3]
            }
        }
    }
}


$users = Get-Content $data -Encoding UTF8 | ForEach-Object {
    if($_ -match '^[a-zA-Z]+$'){
        $user = Get-ADUser $_
        Get-Posts -fio $user.name -users @($user)
    }else{
        $users = Get-ADUser -LDAPFilter "(|(Name=*$_*)(dispalyName=*$_*)(cn=*$_*)(UserPrincipalName=*$_*))" -SearchBase "DC=oek,DC=ru" -Properties *
        Get-Posts -fio $_ -users $users
    }
}

if($csv -gt 0){
    $users | Export-CSV $csv -NoTypeInformation -Delimiter $delimeter -Encoding utf8
}else{
    $users | Format-Table
}