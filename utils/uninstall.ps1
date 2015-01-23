write-host " *** Exchange DEA UNINSTALL Script ***" -f "blue"

net stop MSExchangeTransport 
Disable-TransportAgent -Identity "Exchange DEA" 
Uninstall-TransportAgent -Identity "Exchange DEA" 
net start MSExchangeTransport 
iisreset

write-host "Uninstallation complete. Check previous outputs for any errors!"  -f "yellow"
