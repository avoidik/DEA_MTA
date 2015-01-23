write-host "*** Exchange DEA Install Script ***" -f "blue"

net stop MSExchangeTransport 
Install-TransportAgent -Name "Exchange DEA" -TransportAgentFactory "DEA_MTA.CatchAllFactory" -AssemblyPath "C:\Program Files\Exchange DEA\DEA_MTA.dll"
Set-TransportAgent "Exchange DEA" -Priority 7
Enable-TransportAgent -Identity "Exchange DEA"
net start MSExchangeTransport 

write-host "Installation complete. Check previous outputs for any errors!" -f "yellow" 
