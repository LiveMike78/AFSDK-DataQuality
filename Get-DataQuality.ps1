try {
    # include logging Put-Log $log %message%
    . "C:\Apps\PowerShell\lib\Put-Log.ps1"    

}
catch {
    Write-Host "Error while loading supporting PowerShell Scripts"
    Exit
}

# globals
$log = $MyInvocation.MyCommand.Name
$AFDB = "SmartHome"

[System.Reflection.Assembly]::LoadWithPartialName("OSIsoft.AFSDK") | Out-Null  
$PISystems=New-object OSIsoft.AF.PISystems  
$PISystem=$PISystems.DefaultPISystem  
$myAFDB=$PISystem.Databases[$AFDB]  


$ElementSearch = New-Object OSIsoft.AF.Search.AFElementSearch($myAFDB, "TemplateSearch", "Template:\DataQuality.ByTemplate\")
$ElementList = $ElementSearch.FindObjects()


ForEach ($ele in $ElementList){
    
    $debug = $ele.Attributes.Item("Debug").GetValue().Value.Value

    $ele.GetPath()

    if ($debug -eq 1) {
        Put-Log $log $ele.GetPath()        
    }

    $count = 0
    $stale = 0
    $bad = 0
    $oor = 0

    ForEach ($att in $ele.Attributes) {
        
        if ($null -eq $att.DataReference) {
            $dr = "none"
        }
        else {

            $dr = $att.DataReference.GetType().Name    
        }
        
        # only perform checks on PI Point data
        $notPIPoint = $dr -ne "PIPointDR"

        # exclude PI Point data with Category "Data Quality" (these are the outputs of this script)
        $isDataQuality = $att.Categories.Contains("Data Quality")

        # exclude PI Point summary
        $isSummary = $att.ConfigString.StartsWith("|")

        $skip = $isDataQuality -or $notPIPoint -or $isSummary

        If ($skip -eq $false) {

            # increment the counter to track how many attributes/points are being checked
            $count += 1
            
            $value = ($att.GetValue())

            $isGood = $value.isGood

            #$ST=[datetime]($att.GetValue().Timestamp)
            $ST=[datetime]($value.Timestamp.LocalTime)
            $ET=(Get-Date)

            $age = New-Timespan -Start $ST -End $ET

            $pipt = $att.DataReference.pipoint

            try {
                $pipt.LoadAttributes("ExcMax")
                $max = $pipt.GetAttribute("ExcMax")                    
            }
            catch {
                $max = 600
            }

            try {
                $pipt.LoadAttributes("Zero")
                $zero = $pipt.GetAttribute("Zero")                    
            }
            catch {
                $zero = 0
            }

            try {
                $pipt.LoadAttributes("Span")
                $span = $pipt.GetAttribute("Span")                    
            }
            catch {
                $span = 100
            }

            
            # stale if the snapshot is older than now take the exception maximum x2
            if ($age.TotalSeconds -gt ($max*2)) {
                $stale += 1
                if ($debug -eq 1) {
                    Put-Log $log ($att.GetPath()+": "+$age.TotalSeconds+" > "+$max)
                    ($att.GetPath()+": "+$age.TotalSeconds+" > "+$max)
                }                
            }

            if ($isGood -eq $false) {
                $bad += 1
                if ($debug -eq 1) {
                    Put-Log $log ($att.GetPath()+": "+$att.GetValue()+" = bad")
                    ($att.GetPath()+": "+$att.GetValue()+" = bad")
                }                
            }
            
            
            switch ($value.ValueType.name)
            {
                {($_ -eq "Single") -or ($_ -eq "Double") -or ($_ -eq "Int16") -or ($_ -eq "Int32")}
                {
                    if ($value.value -gt ($zero + $span)) {
                        $oor += 1
                        if ($debug -eq 1) {
                            Put-Log $log ($att.GetPath()+": "+$att.GetValue()+" > "+($zero + $span))
                            ($att.GetPath()+": "+$att.GetPath()+": "+$att.GetValue()+" > "+($zero + $span))
                        }                
                    }
        
                    if ($value.value -lt $zero) {
                        $oor += 1
                        if ($debug -eq 1) {
                            Put-Log $log ($att.GetPath()+": "+$att.GetValue()+" < "+$zero)
                            ($att.GetPath()+": "+$att.GetPath()+": "+$att.GetValue()+" < "+$zero)
                        }                
                    }
                }

                default {
                    #$value.valuetype.name
                }
            }

        }

    }

    If ($count -gt 0) {

        $ele.Attributes.Item("PI Point Count").Data.UpdateValue($count, 0, 1)
        $ele.Attributes.Item("PI Point Stale").Data.UpdateValue($stale, 0, 1)
        $ele.Attributes.Item("PI Point Bad").Data.UpdateValue($bad, 0, 1)
        $ele.Attributes.Item("PI Point Out of Range").Data.UpdateValue($oor, 0, 1)

        $ele.Attributes.Item("Update").Data.UpdateValue($ET, 0, 1)

    }
}
