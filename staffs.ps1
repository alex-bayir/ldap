param(
    [Parameter(Mandatory=$true)][String]$data,
    [string]$csv,
    [string]$delimeter=$null
)

$users = Get-Content $data -Encoding UTF8 | ForEach-Object {
    $body = @{ 'phrase' = $_ }
    $response = Invoke-WebRequest -Method Post 'http://uneco2.ru/local/ajax/header_search.php' -Body $body -ContentType 'application/x-www-form-urlencoded; charset=UTF-8'
    $response.ParsedHtml.getElementsByClassName("popup__employees-list")[0].getElementsByTagName("a") | ForEach-Object{
        $url = "https://uneco2.ru/"+$_.pathname
        $tmp =(Invoke-WebRequest $url)
        $html=$tmp.ParsedHtml
        $arr=$html.getElementsByClassName("arrow-left") | ForEach-Object { $_.children[0].title } | Where-Object { $_ -ne $null -and $_ -ne "" -and $_ -ne "Структура компании" }
        [PSCustomObject]@{
            url = $url
            name = $html.getElementsByClassName("content-item__title")[0].textContent
            post = $html.getElementsByClassName("content-item__text")[0].textContent
            1 = $arr[0]
            2 = $arr[1]
            3 = $arr[2]
            4 = $arr[3]
        }
    }
}

if($csv -gt 0){
    $users | Export-CSV $csv -NoTypeInformation -Delimiter $delimeter -Encoding utf8
}else{
    $users | Format-Table
}