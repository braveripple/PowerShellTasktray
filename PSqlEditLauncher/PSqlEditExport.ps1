$connect_info_reg = Get-ItemProperty HKCU:\Software\OGAWA\POSTGRESQL\PSqlEdit\CONNECT_INFO\ -Name *
$count = $connect_info_reg.COUNT
$connect_info = foreach($i in 0..$($count-1)) {
    $connect_info_reg |
        select-object  `
        @{ Name = "User"; Expression = { $_."USER-$i" } }, `
        @{ Name = "Password"; Expression = { $_."PASSWD-$i" } }, `
        @{ Name = "DBName"; Expression = { $_."DBNAME-$i" } }, `
        @{ Name = "Host"; Expression = { $_."HOST-$i" } }, `
        @{ Name = "PortNo"; Expression = { $_."PORT-$i" } }, `
        @{ Name = "Memo"; Expression = { $_."CONNECT-NAME-$i" } }
}
$connect_info | Export-Csv -LiteralPath "接続先_$(Get-Date -Format FileDateTime).tsv" -Delimiter "`t" -UseQuotes Never
