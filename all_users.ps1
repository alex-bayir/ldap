param(
    [String]$out="all_users.csv",
    [string]$delimiter=";",
    [string]$server,
    [string]$ADCModule="Y:\Corps\СИБ\_ДИБ\Личные папки\Байрашный\ActiveDirectory"
)

if (!(Get-Module -All -Name *ActiveDirectory*)) {
    Import-Module $ADCModule"\Microsoft.ActiveDirectory.Management.resources.dll"
    Import-Module $ADCModule"\Microsoft.ActiveDirectory.Management.dll"
    Import-Module $ADCModule"\ActiveDirectory.psd1"
}

function Fix-Name($name) {
    return $(if($name -is [String]) { $name.trim() -replace "  "," " } else { $null })
}
if($server){
    $users=Get-ADUser -Server $server -Filter * -SearchBase "DC=oek,DC=ru" -Properties SamAccountName,Enabled,CanonicalName,EmailAddress,Name,EmployeeID,Title,MemberOf
}else{
    $users=Get-ADUser -Filter * -SearchBase "DC=oek,DC=ru" -Properties SamAccountName,Enabled,CanonicalName,EmailAddress,Name,EmployeeID,Title,MemberOf
}

$list=$users | ForEach-Object {
    $canonical=if($_){($_.CanonicalName -replace "/$($_.Name)","").split("/")}else{@()}
    return [PSCustomObject]@{
        account = $_.SamAccountName
        enabled = $_.Enabled
        email = $_.EmailAddress
        name = Fix-Name $_.Name
        employee = $_.EmployeeID
        post = Fix-Name $_.Title
        office = Fix-Name $canonical[2]
        departament = Fix-Name $canonical[3]
        division = Fix-Name $canonical[4]
        domain = $canonical[0]
        root = $canonical[1]
        groups = ($_.MemberOf | ForEach-Object { ($_ -replace "(CN|OU|DC)=","") -replace ",","/" }) -join "`n"
    }
} | Sort-Object -Property account
$list | Export-CSV $out -NoTypeInformation -Delimiter $delimiter -Encoding utf8
