Voir si l’on peux sinon passer un Script Custom qui récupère les même info
Comme par exemple en powershell : Get-NetTCPConnection |  select-object LocalAddress,LocalPort,RemoteAddress,RemotePort,State,CreationTime,OwningProcess, @{Name="Process";Expression={(Get-Process -Id $_.OwningProcess).ProcessName}} | ft -AutoSize (pour state=Listen suffisant, je ne sais pas comment restreindre)
