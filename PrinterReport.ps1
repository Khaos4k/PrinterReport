$printerlist = Import-Csv c:\scripts\input\printerlist.csv
$SNMP = New-Object -ComObject olePrn.OleSNMP
$PrinterStats = @()

foreach ( $p in $printerlist)
{
    $tonervolume = $null
    $currentvolume = $null
    $percentremaining = $null
    $statustree = $null
    $status = $null
    $toner = $null
    
    if (Test-Connection -ComputerName $p.ip -Count 1 -quiet)
    {
        $snmp.open($p.ip,"public",2,3000)
        #$printertype = $snmp.Get(".1.3.6.1.2.1.25.3.2.1.3.1")
        $model = $p.model
        
        if ($model -like "*PRINTER NAME*")
        {        
            $tonervolume = $snmp.get("43.11.1.1.8.1.1")
            $currentvolume = $snmp.get("43.11.1.1.9.1.1")
            [int]$percentremaining = ($currentvolume / $tonervolume) * 100
            $toner = $percentremaining.ToString() + ' %'
     
            $statustree = $snmp.gettree("43.18.1.1.8")
            $status = $statustree|? {$_ -notlike "print*"} #status, including low ink warnings
            $status = $status|? {$_ -notlike "*bypass*"}

            $PrinterStats += New-Object -TypeName PSObject -Property @{
                Location = $p.location
                "Printer Name" = $p.model
                "IP Address" = $p.ip
                "Toner Remaining" = $toner
                Status = $status
            }
        }
    }
    if (!(Test-Connection -ComputerName $p.ip -Count 1 -quiet))
    {
        $PrinterStats += New-Object -TypeName PSObject -Property @{
            Location = $p.location
            "Printer Name" = $p.model
            "IP Address" = $p.ip
            "Toner Remaining" = $toner
            Status = "OFFLINE"
        }
    }

    
} $PrinterStats
