Set UpdateSession = CreateObject("Microsoft.Update.Session")
Set UpdateServiceManager = CreateObject("Microsoft.Update.ServiceManager")
Set UpdateService = UpdateServiceManager.AddScanPackageService("Offline Sync Service", "c:\scripts\wsusscn2.cab", 1)
Set UpdateSearcher = UpdateSession.CreateUpdateSearcher()

WScript.Echo "Searching for updates..." & vbCRLF

UpdateSearcher.ServerSelection = 3 ' ssOthers

UpdateSearcher.IncludePotentiallySupersededUpdates = False 'good for older OSes, to include Security-Only or superseded updates in the result list, otherwise these are pruned out and not returned as part of the final result list

UpdateSearcher.ServiceID = UpdateService.ServiceID

Set SearchResult = UpdateSearcher.Search("IsInstalled=0") 'or "IsInstalled=0 or IsInstalled=1" to also list the installed updates as MBSA did

Set Updates = SearchResult.Updates

If searchResult.Updates.Count = 0 Then
    WScript.Echo "There are no applicable updates."
    WScript.Quit
End If

WScript.Echo "List of applicable items on the machine when using wssuscan.cab:" & vbCRLF

For I = 0 to searchResult.Updates.Count-1
    Set update = searchResult.Updates.Item(I)
    WScript.Echo I + 1 & "> " & update.Title
Next

WScript.Quit