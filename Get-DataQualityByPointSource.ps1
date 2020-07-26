# destination for CSV output
$reportFile = "C:\apps\PowerShell\dat\dq.csv"

[System.Reflection.Assembly]::LoadWithPartialName("OSIsoft.AFSDK") | Out-Null

# get default PI Data Archive
$piSrvs = New-Object OSIsoft.AF.PI.PIServers  
$piSrv = $piSrvs.DefaultPIServer  

# list all point sources
$ptSources = $piSrv.PointSources

# create CSV file with headers if it doesn't exit
if ([system.io.file]::exists($reportFile) -eq $False){
    "DateTime,Point Source,BadOnly,StaleOnly,BadStale,Good,Total" | Out-File -filepath $reportFile
}
 

#$ET = (Get-Date).AddDays(-90).AddHours(-1)

#While($ET -lt (Get-Date)) {
    

# process each point source
ForEach ($ptSource in $ptSources) {
    
    # find all points with the point source
    $piPointQuery = New-Object OSIsoft.AF.PI.PIPointQuery([OSIsoft.AF.PI.PICommonPointAttributes]::PointSource, [OSIsoft.AF.Search.AFSearchOperator]::Equal, $ptSource.Name)
    $queryList = New-Object System.Collections.Generic.List[OSIsoft.AF.PI.PIPointQuery]
    $queryList.Add($piPointQuery)

    [Type[]]$types = [OSIsoft.AF.PI.PIServer], [System.Collections.Generic.IEnumerable[OSIsoft.AF.PI.PIPointQuery]], [System.Collections.Generic.IEnumerable[string]]  

    $FindPIPointMethod = [OSIsoft.AF.PI.PIPoint].GetMethod("FindPIPoints",$types)  
    $Points = $FindPIPointMethod.Invoke($null,@($piSrv,[System.Collections.Generic.List[OSIsoft.AF.PI.PIPointQuery]]$queryList,$null))   

    $PointList = New-Object OSISoft.AF.PI.PIPointList -ArgumentList $Points

    # set the start time and end time for the query
     $ET=(Get-Date)
    #$ET = $ET.AddHours(1)
    
    $TimeStale = $ET.AddHours(-8)
    
    # find all points stale (no updates in query time), bad (quality) and stale & bad in combination
    $StaleList =  $PointList.CurrentValue() | Where-Object {$_.Timestamp -le $TimeStale -and $_.Status -ne 'Bad'}
    $BadList = $PointList.CurrentValue() | Where-Object {$_.Status -eq 'Bad' -and $_.Timestamp -gt $TimeStale}
    $BadStaleList = $PointList.CurrentValue() | Where-Object {$_.Status -eq 'Bad' -and $_.Timestamp -le $TimeStale}

    # calculate the "good" points
    $GoodCount = $PointList.Count - $StaleList.Count - $BadList.Count - $BadStaleList.Count

    # append to the CSV file
    ($ET.ToUniversalTime().toString("yyyy-MM-ddTHH:mm:ssZ")+","+$ptSource.Name+","+$BadList.Count+","+$StaleList.Count+","+$BadStaleList.Count+","+$GoodCount+","+$PointList.Count) | Out-File -filepath $reportFile -Append 
    
}

#}