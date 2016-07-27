$printerlist = Import-Csv x:\xtyler\scripts\import\printerlist.csv
$SNMP = New-Object -ComObject olePrn.OleSNMP

foreach ( $p in $printerlist)
{
    if (Test-Connection -ComputerName $p.ip -Count 1 -Quiet)
    {
        $snmp.open($p.value,"public",2,3000)
        $printertype = $snmp.Get(".1.3.6.1.2.1.25.3.2.1.3.1")
    }
    if (!(Test-Connection -ComputerName $p.ip -Count 1 -Quiet))
    {
        Write ($p.location +" is offline")
    }

    if ($printertype -like "*WorkCentre 5755*")
    {        
        $tonervolume = $snmp.get("43.11.1.1.8.1.1")
        $currentvolume = $snmp.get("43.11.1.1.9.1.1")
        [int]$percentremaining = ($currentvolume / $tonervolume) * 100
 
        $statustree = $snmp.gettree("43.18.1.1.8")
        $status = $statustree|? {$_ -notlike "print*"} #status, including low ink warnings
        $status = $status|? {$_ -notlike "*bypass*"}

        $PrinterStats = New-Object -TypeName PSObject -property @{
            tonervolume = $tonervolume
            currentvolume = $currentvolume
            "Printer Name" = $p.name
            "IP Address" = $p.ip
            Location = $p.location
            "Toner Remaining" = $percentremaining
            Status = $status
        }
    }
}

