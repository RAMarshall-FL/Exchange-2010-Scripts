Function Get-ClientConnections
{
    # List of AD Sites with Exchange Client Access Servers
    Write-Host "`n"
    $AdSiteList = Get-ExchangeServer | Select-Object Site -Unique
    $AdSiteList = $AdSiteList -replace ".*sites/" -replace "" -replace "}" -replace ""
    $AdSiteMenu = @{}

    Write-Host `n ; Write-Host ("*"*(52))
    Write-Host "** " -NoNewline ; Write-Host "Active Directory Site(s) with Exchange Servers" -ForegroundColor Magenta -NoNewline ; Write-Host " **"
    Write-Host ("*"*(52))`n

    # Build AD site menu
    For ($i=1; $i -le ($AdSiteList | Measure-Object).Count ; $i++)
    {
        Write-Host "$i.$($AdSiteList[$i-1])"
        $AdSiteMenu.add($i,($AdSiteList[$i-1]))
    } # End For Loop
    
    Do
    {
        Write-Host "`nCheck Client Connections for Exchange Servers in which AD Site: " -ForegroundColor Yellow -NoNewline
        [int]$AdSiteChoice = Read-Host
        $AdSite = $AdSiteMenu.Item($AdSiteChoice)
    }
    Until ($AdSiteMenu.item($AdSiteChoice) -ne $Null)

    # Get Exchange Client Access Servers in selected AD site
    Try
    {
        $ClientAccessServers = Get-ExchangeServer -Status | Where-Object {$_.site -match $AdSite -AND $_.serverRole -match "ClientAccess"} | Sort-Object Name -ErrorAction Stop
        Write-Host "`nChecking Client Connections of Exchange Servers in AD Site: " -ForegroundColor White -NoNewline
        Write-Host "$AdSite`n" -ForegroundColor Cyan
    }
    Catch
    {
        Write-Host "No Client Access Servers found in AD site: " -ForegroundColor Red -NoNewline
        Write-Host $AdSite
    }

    # Select connection type to check
    Do
    {
        Write-Host "Check Connections for: OWA (1), RPC (2), ActiveSync (3), or All (4): " -ForegroundColor Yellow -NoNewline
        [int]$ConnectionType = Read-Host
        Write-Host `n
    }
    Until (($ConnectionType) -ne $Null)

    Foreach ($CAS in $ClientAccessServers)
    {
        If ($ConnectionType -eq "1")
        {
            # Get OWA Connections
            $OwaValue = Get-Counter -ComputerName $CAS.name -Counter "\MsExchange OWA\Current Unique Users" | Foreach {$_.counterSamples[0].cookedValue}
            Write-Host "`Total OWA Connections on server: " -NoNewline
            Write-Host $CAS.name -ForegroundColor Cyan -NoNewline ; Write-Host "   $OwaValue`n" -ForegroundColor Green
        }
        ElseIf ($ConnectionType -eq "2")
        {
            # Get RPC Connections
            $RpcValue = Get-Counter -ComputerName $CAS.name -Counter "\MsExchange RpcClientAccess\User Count" | Foreach {$_.counterSamples[0].cookedValue}
            Write-Host "`Total RPC Connections on server: " -NoNewline
            Write-Host $CAS.name -ForegroundColor Cyan -NoNewline ; Write-Host "   $RpcValue`n" -ForegroundColor Green
        }
        ElseIf ($ConnectionType -eq "3")
        {
            # Get Async Connections
            $ASyncValue = Get-Counter -ComputerName $CAS.name -Counter "\MsExchange ActiveSync\Current Requests" | Foreach {$_.counterSamples[0].cookedValue}
            Write-Host "`Total Active Sync Connections on server: " -NoNewline
            Write-Host $CAS.name -ForegroundColor Cyan -NoNewline ; Write-Host "   $ASyncValue`n" -ForegroundColor Green
        }
        ElseIf ($ConnectionType -eq "4")
        {
            # Get OWA Connections
            $OwaValue = Get-Counter -ComputerName $CAS.name -Counter "\MsExchange OWA\Current Unique Users" | Foreach {$_.counterSamples[0].cookedValue}
            Write-Host "`Total OWA Connections on server: " -NoNewline
            Write-Host $CAS.name -ForegroundColor Cyan -NoNewline ; Write-Host "   $OwaValue`n" -ForegroundColor Green

            # Get RPC Connections
            $RpcValue = Get-Counter -ComputerName $CAS.name -Counter "\MsExchange RpcClientAccess\User Count" | Foreach {$_.counterSamples[0].cookedValue}
            Write-Host "`Total RPC Connections on server: " -NoNewline
            Write-Host $CAS.name -ForegroundColor Cyan -NoNewline ; Write-Host "   $RpcValue`n" -ForegroundColor Green

            # Get Async Connections
            $ASyncValue = Get-Counter -ComputerName $CAS.name -Counter "\MsExchange ActiveSync\Current Requests" | Foreach {$_.counterSamples[0].cookedValue}
            Write-Host "`Total Active Sync Connections on server: " -NoNewline
            Write-Host $CAS.name -ForegroundColor Cyan -NoNewline ; Write-Host "   $ASyncValue`n" -ForegroundColor Green
        }
    } # End Foreach 
}