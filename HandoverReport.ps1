<#
--------------------------------------------------------------------------------------------------------
Name: PagerDuty Incident Report
Author: Ayush Lal
Description: PowerShell script that fetches PagerDuty Incident details via their API and then generates 
a HTML file that can be emailed.
--------------------------------------------------------------------------------------------------------
#>

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$apiKey = "API KEY"

# URI for Triggered PD alerts for a specifc Team
$URI_triggered = "https://api.pagerduty.com/incidents?statuses[]=triggered&team_ids[]=TEAM_ID"

# URI for Acknowledged PD alerts for a specifc Team
$URI_ack = "https://api.pagerduty.com/incidents?statuses[]=acknowledged&team_ids[]=TEAM_ID"

$headers = @{
    'Authorization' = $apiKey
    'Content-type'  = 'application/json'
    'Accept'        = 'application/vnd.pagerduty+json;version=2'
}

# Returns Triggered Incidents
$response = Invoke-RestMethod -Method Get -Uri $URI_triggered -Header $headers
$results = $response.incidents | Select-Object summary, status, html_url, created_at

# Returns Acknowledged Incidents
$response2 = Invoke-RestMethod -Method GET -Uri $URI_ack -Header $headers
$results2 = $response2.incidents | Select-Object summary, priority, status, last_status_change_by, created_at, html_url


# Variables for both $results and $results2 objects to be used for the HTML table formatting
# Data for $results
$rSummary = $results.summary
$rStatus = $results.status
$rUrl = $results.html_url

if ($results) {
    Write-Host "Showing Triggered alerts:" -ForegroundColor Green
    $rSummary + ", " + $rStatus + ", " + $r2DateTimeResult + ", " + $rUrl
}
else {
    Write-Host "There are no Triggered alerts" -ForegroundColor Magenta
}

# Data for $results2
$r2Summary = $results2.summary 
$r2Priority = $results2.priority.summary
$r2Status = $results2.status
$r2Url = $results2.html_url
$r2Changeby = $results2.last_status_change_by.summary

if ($results2) {
    Write-Host "Showing Acknowledged alerts:" -ForegroundColor Green
    $r2Summary + ", " + $r2Priority + ", " + $r2Status + ", " + $r2Url + ", " + $r2Changeby + ", " + $rDateTimeResult | Format-Table -auto
}
else {
    Write-Host "There are no Acknowledged alerts" -ForegroundColor Magenta
}

$ack_inc_count = ($r2Summary.count)
$triggered_inc_count = ($rSummary.count)

Write-Host "There are currently $ack_inc_count Acknowledged and $triggered_inc_count Triggered alerts for the NOC." -ForegroundColor green

Write-Host ""
Write-Host ""
Write-Host ""

# Building Array for $results2 (Acknowledged Incidents)
# Array response variables for $results2.summary
$r2Array_sum0 = $r2Summary # used for array if there is only one item...
if ($r2Summary.count -gt 0) {
    $r2Array_sum1 = $r2Summary[0]
    $r2Array_sum2 = $r2Summary[1]
    $r2Array_sum3 = $r2Summary[2]
    $r2Array_sum4 = $r2Summary[3]
    $r2Array_sum5 = $r2Summary[4]
}

# Array response variables for $results2.status
$r2Array_status0 = $r2Status
if ($r2Status.count -gt 0) {
    $r2Array_status1 = $r2Status[0]
    $r2Array_status2 = $r2Status[1]
    $r2Array_status3 = $r2Status[2]
    $r2Array_status4 = $r2Status[3]
    $r2Array_status5 = $r2Status[4]
}

# Array response variables for $results2.html_url
$r2Array_URL0 = $r2URL 
if ($r2Url.count -gt 0) {
    $r2Array_URL1 = $r2Url[0]
    $r2Array_URL2 = $r2Url[1]
    $r2Array_URL3 = $r2Url[2]
    $r2Array_URL4 = $r2Url[3]
    $r2Array_URL5 = $r2Url[4]
}

# Array response variables for $results2.last_status_change_by
$r2Array_Changeby0 = $r2Changeby
if ($r2Changeby.count -gt 0) {
    $r2Array_Changeby1 = $r2Changeby[0]
    $r2Array_Changeby2 = $r2Changeby[1]
    $r2Array_Changeby3 = $r2Changeby[2]
    $r2Array_Changeby4 = $r2Changeby[3]
    $r2Array_Changeby5 = $r2Changeby[4]
}

# Building Array for $results (Triggered Incidents)
# Array response variables for $results.summary
$rArray_sum0 = $rSummary
if ($rSummary.count -gt 0) {
    $rArray_sum1 = $rSummary[0]
    $rArray_sum2 = $rSummary[1]
    $rArray_sum3 = $rSummary[2]
    $rArray_sum4 = $rSummary[3]
    $rArray_sum5 = $rSummary[4]
}

# Array response variables for $results.status
$rArray_status0 = $rStatus
if ($rStatus.count -gt 0) {
    $rArray_status1 = $rStatus[0]
    $rArray_status2 = $rStatus[1]
    $rArray_status3 = $rStatus[2]
    $rArray_status4 = $rStatus[3]
    $rArray_status5 = $rStatus[4]
}

# Array response variables for $results.url (Triggered)
$rArray_URL0 = $rUrl
if ($rUrl.count -gt 0) {
    $rArray_URL1 = $rUrl[0]
    $rArray_URL2 = $rUrl[1]
    $rArray_URL3 = $rUrl[2]
    $rArray_URL4 = $rUrl[3]
    $rArray_URL5 = $rUrl[4]
}

# Start compiling the Multidemensional Array Tables
Write-Host "Building Array Table for Triggered Incidents:"
$emp_counter = $null  
$Triggered_Table = @()

$emp_counter ++
if ($rSummary.count -eq 1) {
    $T_Table1 = $Triggered_Table += , @($emp_counter, $rArray_sum0, $rArray_status0, $rArray_URL0)

    if ($T_Table1) {
        $T_HTML_Table += '<table>'
        $T_HTML_Table += '<tr>'
        $T_HTML_Table += "<th class='Tb_heading' colspan='3'>PagerDuty Triggered Alerts</th>"
        $T_HTML_Table += '</<tr>'
        $T_HTML_Table += '<tr>'
        $T_HTML_Table += '<th>Alert Summary</th>'
        $T_HTML_Table += '<th>Status</th>'
        $T_HTML_Table += '<th>URL</th>'
        $T_HTML_Table += '</tr>'

        $T_HTML_Table += '<tr>'
        $T_HTML_Table += "<td><p>$rArray_sum0 </p></td>"
        $T_HTML_Table += "<td><p>$rArray_status0</p></td>"
        $T_HTML_Table += "<td><a href='$rArray_URL0'>$rArray_URL0</p></td>"
        $T_HTML_Table += '</tr>'
        $T_HTML_Table += '</table>'
        $T_HTML_Table += '<br>'
    }
    else {
        Write-Host "not working"
    }
}
elseif ($rSummary.count -eq 2) {
    $T_Table2 = $Triggered_Table += , @($emp_counter, $rArray_sum1, $rArray_status1, $rArray_URL1)
    $emp_counter ++
    $T_Table2 = $Triggered_Table += , @($emp_counter, $rArray_sum2, $rArray_status2, $rArray_URL2)

    if ($T_Table2) {
        $T_HTML_Table += '<table>'
        $T_HTML_Table += '<tr>'
        $T_HTML_Table += "<th class='Tb_heading' colspan='3'>PagerDuty Triggered Alerts</th>"
        $T_HTML_Table += '</<tr>'
        $T_HTML_Table += '<tr>'
        $T_HTML_Table += '<th>Alert Summary</th>'
        $T_HTML_Table += '<th>Status</th>'
        $T_HTML_Table += '<th>URL</th>'
        $T_HTML_Table += '</tr>'

        $T_HTML_Table += '<tr>'
        $T_HTML_Table += "<td><p>$rArray_sum1</p></td>"
        $T_HTML_Table += "<td><p>$rArray_status1</p></td>"
        $T_HTML_Table += "<td><a href='$rArray_URL1'>$rArray_URL1</p></td>"
        $T_HTML_Table += '</tr>'

        $T_HTML_Table += '<tr>'
        $T_HTML_Table += "<td><p>$rArray_sum2</p></td>"
        $T_HTML_Table += "<td><p>$rArray_status2</p></td>"
        $T_HTML_Table += "<td><a href='$rArray_URL2'>$rArray_URL2</p></td>"
        $T_HTML_Table += '</tr>'
        $T_HTML_Table += '</table>'
        $T_HTML_Table += '<br>'
    }
    else {
        Write-Host "not working"
    }
}
elseif ($rSummary.count -eq 3) {
    $T_Table3 = $Triggered_Table += , @($emp_counter, $rArray_sum1, $rArray_status1, $rArray_URL1)
    $emp_counter ++
    $T_Table3 = $Triggered_Table += , @($emp_counter, $rArray_sum2, $rArray_status2, $rArray_URL2)
    $emp_counter ++  
    $T_Table3 = $Triggered_Table += , @($emp_counter, $rArray_sum3, $rArray_status3, $rArray_URL3)

    if ($T_Table3) {
        $T_HTML_Table += '<table>'
        $T_HTML_Table += '<tr>'
        $T_HTML_Table += "<th class='Tb_heading' colspan='3'>PagerDuty Triggered Alerts</th>"
        $T_HTML_Table += '</<tr>'
        $T_HTML_Table += '<tr>'
        $T_HTML_Table += '<th>Alert Summary</th>'
        $T_HTML_Table += '<th>Status</th>'
        $T_HTML_Table += '<th>URL</th>'
        $T_HTML_Table += '</tr>'

        $T_HTML_Table += '<tr>'
        $T_HTML_Table += "<td><p>$rArray_sum1</p></td>"
        $T_HTML_Table += "<td><p>$rArray_status1</p></td>"
        $T_HTML_Table += "<td><a href='$rArray_URL1'>$rArray_URL1</p></td>"
        $T_HTML_Table += '</tr>'

        $T_HTML_Table += '<tr>'
        $T_HTML_Table += "<td><p>$rArray_sum2</p></td>"
        $T_HTML_Table += "<td><p>$rArray_status2</p></td>"
        $T_HTML_Table += "<td><a href='$rArray_URL2'>$rArray_URL2</p></td>"
        $T_HTML_Table += '</tr>'

        $T_HTML_Table += '<tr>'
        $T_HTML_Table += "<td><p>$rArray_sum3</p></td>"
        $T_HTML_Table += "<td><p>$rArray_status3</p></td>"
        $T_HTML_Table += "<td><a href='$rArray_URL3'>$rArray_URL3</p></td>"
        $T_HTML_Table += '</tr>'
        $T_HTML_Table += '</table>'
        $T_HTML_Table += '<br>'
    }
    else {
        Write-Host "Not working"
    }
}
elseif ($rSummary.count -eq 4) {
    $T_Table4 = $Triggered_Table += , @($emp_counter, $rArray_sum1, $rArray_status1, $rArray_URL1)
    $emp_counter ++
    $T_Table4 = $Triggered_Table += , @($emp_counter, $rArray_sum2, $rArray_status2, $rArray_URL2)
    $emp_counter ++  
    $T_Table4 = $Triggered_Table += , @($emp_counter, $rArray_sum3, $rArray_status3, $rArray_URL3)
    $emp_counter ++  
    $T_Table4 = $Triggered_Table += , @($emp_counter, $rArray_sum4, $rArray_status4, $rArray_URL4)

    if ($T_Table4) {
        $T_HTML_Table += '<table>'
        $T_HTML_Table += '<tr>'
        $T_HTML_Table += "<th class='Tb_heading' colspan='3'>PagerDuty Triggered Alerts</th>"
        $T_HTML_Table += '</<tr>'
        $T_HTML_Table += '<tr>'
        $T_HTML_Table += '<th>Alert Summary</th>'
        $T_HTML_Table += '<th>Status</th>'
        $T_HTML_Table += '<th>URL</th>'
        $T_HTML_Table += '</tr>'

        $T_HTML_Table += '<tr>'
        $T_HTML_Table += "<td><p>$rArray_sum1</p></td>"
        $T_HTML_Table += "<td><p>$rArray_status1</p></td>"
        $T_HTML_Table += "<td><a href='$rArray_URL1'>$rArray_URL1</p></td>"
        $T_HTML_Table += '</tr>'

        $T_HTML_Table += '<tr>'
        $T_HTML_Table += "<td><p>$rArray_sum2</p></td>"
        $T_HTML_Table += "<td><p>$rArray_status2</p></td>"
        $T_HTML_Table += "<td><a href='$rArray_URL2'>$rArray_URL2</p></td>"
        $T_HTML_Table += '</tr>'

        $T_HTML_Table += '<tr>'
        $T_HTML_Table += "<td><p>$rArray_sum3</p></td>"
        $T_HTML_Table += "<td><p>$rArray_status3</p></td>"
        $T_HTML_Table += "<td><a href='$rArray_URL3'>$rArray_URL3</p></td>"
        $T_HTML_Table += '</tr>'

        $T_HTML_Table += '<tr>'
        $T_HTML_Table += "<td><p>$rArray_sum4</p></td>"
        $T_HTML_Table += "<td><p>$rArray_status4</p></td>"
        $T_HTML_Table += "<td><a href='$rArray_URL4'>$rArray_URL4</p></td>"
        $T_HTML_Table += '</tr>'
        $T_HTML_Table += '</table>'
        $T_HTML_Table += '<br>'
    }
    else {
        Write-Host "Not working"
    }
}
else {
    $T_Table5 = $Triggered_Table += , @($emp_counter, $rArray_sum1, $rArray_status1, $rArray_URL1)
    $emp_counter ++  
    $T_Table5 = $Triggered_Table += , @($emp_counter, $rArray_sum2, $rArray_status2, $rArray_URL2)  
    $emp_counter ++  
    $T_Table5 = $Triggered_Table += , @($emp_counter, $rArray_sum3, $rArray_status3, $rArray_URL3)  
    $emp_counter ++  
    $T_Table5 = $Triggered_Table += , @($emp_counter, $rArray_sum4, $rArray_status4, $rArray_URL4) 
    $emp_counter ++  
    $T_Table5 = $Triggered_Table += , @($emp_counter, $rArray_sum5, $rArray_status5, $rArray_URL5)

    if ($T_Table5) {
        $T_HTML_Table += '<table>'
        $T_HTML_Table += '<tr>'
        $T_HTML_Table += "<th class='Tb_heading' colspan='3'>PagerDuty Triggered Alerts</th>"
        $T_HTML_Table += '</<tr>'
        $T_HTML_Table += '<tr>'
        $T_HTML_Table += '<th>Alert Summary</th>'
        $T_HTML_Table += '<th>Status</th>'
        $T_HTML_Table += '<th>URL</th>'
        $T_HTML_Table += '</tr>'
        
        $T_HTML_Table += '<tr>'
        $T_HTML_Table += "<td><p>$rArray_sum1</p></td>"
        $T_HTML_Table += "<td><p>$rArray_status1</p></td>"
        $T_HTML_Table += "<td><a href='$rArray_URL1'>$rArray_URL1</p></td>"
        $T_HTML_Table += '</tr>'

        $T_HTML_Table += '<tr>'
        $T_HTML_Table += "<td><p>$rArray_sum2</p></td>"
        $T_HTML_Table += "<td><p>$rArray_status2</p></td>"
        $T_HTML_Table += "<td><a href='$rArray_URL2'>$rArray_URL2</p></td>"
        $T_HTML_Table += '</tr>'

        $T_HTML_Table += '<tr>'
        $T_HTML_Table += "<td><p>$rArray_sum3</p></td>"
        $T_HTML_Table += "<td><p>$rArray_status3</p></td>"
        $T_HTML_Table += "<td><a href='$rArray_URL3'>$rArray_URL3</p></td>"
        $T_HTML_Table += '</tr>'

        $T_HTML_Table += '<tr>'
        $T_HTML_Table += "<td><p>$rArray_sum4</p></td>"
        $T_HTML_Table += "<td><p>$rArray_status4</p></td>"
        $T_HTML_Table += "<td><a href='$rArray_URL4'>$rArray_URL4</p></td>"
        $T_HTML_Table += '</tr>'

        $T_HTML_Table += '<tr>'
        $T_HTML_Table += "<td><p>$rArray_sum5</p></td>"
        $T_HTML_Table += "<td><p>$rArray_status5</p></td>"
        $T_HTML_Table += "<td><a href='$rArray_URL5'>$rArray_URL5</p></td>"
        $T_HTML_Table += '</tr>'
        $T_HTML_Table += '</table>'
        $T_HTML_Table += '<br>'
    }
    else {
        Write-Host "Not working"
    }
}
write-host "=============================================================="
foreach ($triggered_item in $Triggered_Table) {  
    Write-host ($triggered_item)
}
$Triggered_Table | % { $_ -join ',' } | Out-File -FilePath ".\Triggered_Table.txt"


Write-Host ""
Write-Host "....................................................................................."
Write-Host ""


Write-Host "Building Array Table for Acknowledged Incidents:"
$counter = $null   
$Acknowledged_Table = @()   

$counter ++
if ($r2Summary.count -eq 1) {
    $A_Table1 = $Acknowledged_Table += , @($counter, $r2Array_sum0, $r2Array_status0, $r2Array_Changeby0, $r2Array_URL0)

    if ($A_Table1) {
        $A_HTML_Table += '<table>'
        $A_HTML_Table += '<tr>'
        $A_HTML_Table += "<th class='Tb_heading' colspan='4'>PagerDuty Acknowledged Alerts</th>"
        $A_HTML_Table += '</<tr>'
        $A_HTML_Table += '<tr>'
        $A_HTML_Table += '<th>Alert Summary</th>'
        $A_HTML_Table += '<th>Status</th>'
        $A_HTML_Table += '<th>Last Changed by</th>'
        $A_HTML_Table += '<th>URL</th>'
        $A_HTML_Table += '</tr>'

        $A_HTML_Table += '<tr>'
        $A_HTML_Table += "<td><p>$r2Array_sum0</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_status0</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_Changeby0</p></td>"
        $A_HTML_Table += "<td><a href='$r2Array_URL0'>$r2Array_URL0</p></td>"
        $A_HTML_Table += '</tr>'
        $A_HTML_Table += '</table>'
    }
    else {
        Write-Host "not working"
    }
}
elseif ($r2Summary.count -eq 2) {
    $A_Table2 = $Acknowledged_Table += , @($counter, $r2Array_sum1, $r2Array_status1, $r2Array_Changeby1, $r2Array_URL1)
    $counter ++
    $A_Table2 = $Acknowledged_Table += , @($counter, $r2Array_sum2, $r2Array_status2, $r2Array_Changeby2, $r2Array_URL2)

    if ($A_Table2) {
        $A_HTML_Table += '<table>'
        $A_HTML_Table += '<tr>'
        $A_HTML_Table += "<th class='Tb_heading' colspan='4'>PagerDuty Acknowledged Alerts</th>"
        $A_HTML_Table += '</<tr>'
        $A_HTML_Table += '<tr>'
        $A_HTML_Table += '<th>Alert Summary</th>'
        $A_HTML_Table += '<th>Status</th>'
        $A_HTML_Table += '<th>Last Changed by</th>'
        $A_HTML_Table += '<th>URL</th>'
        $A_HTML_Table += '</tr>'

        $A_HTML_Table += '<tr>'
        $A_HTML_Table += "<td><p>$r2Array_sum1</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_status1</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_Changeby1</p></td>"
        $A_HTML_Table += "<td><a href='$r2Array_URL1'>$r2Array_URL1</p></td>"
        $A_HTML_Table += '</tr>'

        $A_HTML_Table += '<tr>'
        $A_HTML_Table += "<td><p>$r2Array_sum2</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_status2</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_Changeby2</p></td>"
        $A_HTML_Table += "<td><a href='$r2Array_URL2'>$r2Array_URL2</p></td>"
        $A_HTML_Table += '</tr>'
        $A_HTML_Table += '</table>'
    }
    else {
        Write-Host "not working"
    }
}
elseif ($r2Summary.count -eq 3) {
    $A_Table3 = $Acknowledged_Table += , @($counter, $r2Array_sum1, $r2Array_status1, $r2Array_Changeby1, $r2Array_URL1)
    $counter ++
    $A_Table3 = $Acknowledged_Table += , @($counter, $r2Array_sum2, $r2Array_status2, $r2Array_Changeby2, $r2Array_URL2)
    $counter ++  
    $A_Table3 = $Acknowledged_Table += , @($counter, $r2Array_sum3, $r2Array_status3, $r2Array_Changeby3, $r2Array_URL3)

    if ($A_Table3) {
        $A_HTML_Table += '<table>'
        $A_HTML_Table += '<tr>'
        $A_HTML_Table += "<th class='Tb_heading' colspan='4'>PagerDuty Acknowledged Alerts</th>"
        $A_HTML_Table += '</<tr>'
        $A_HTML_Table += '<tr>'
        $A_HTML_Table += '<th>Alert Summary</th>'
        $A_HTML_Table += '<th>Status</th>'
        $A_HTML_Table += '<th>Last Changed by</th>'
        $A_HTML_Table += '<th>URL</th>'
        $A_HTML_Table += '</tr>'

        $A_HTML_Table += '<tr>'
        $A_HTML_Table += "<td><p>$r2Array_sum1</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_status1</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_Changeby1</p></td>"
        $A_HTML_Table += "<td><a href='$r2Array_URL1'>$r2Array_URL1</p></td>"
        $A_HTML_Table += '</tr>'

        $A_HTML_Table += '<tr>'
        $A_HTML_Table += "<td><p>$r2Array_sum2</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_status2</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_Changeby2</p></td>"
        $A_HTML_Table += "<td><a href='$r2Array_URL2'>$r2Array_URL2</p></td>"
        $A_HTML_Table += '</tr>'

        $A_HTML_Table += '<tr>'
        $A_HTML_Table += "<td><p>$r2Array_sum3</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_status3</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_Changeby3</p></td>"
        $A_HTML_Table += "<td><a href='$r2Array_URL3'>$r2Array_URL3</p></td>"
        $A_HTML_Table += '</tr>'
        $A_HTML_Table += '</table>'
    }
    else {
        Write-Host "Not working"
    }
}
elseif ($r2Summary.count -eq 4) {
    $A_Table4 = $Acknowledged_Table += , @($counter, $r2Array_sum1, $r2Array_status1, $r2Array_Changeby1, $r2Array_URL1)
    $counter ++
    $A_Table4 = $Acknowledged_Table += , @($counter, $r2Array_sum2, $r2Array_status2, $r2Array_Changeby2, $r2Array_URL2)
    $counter ++  
    $A_Table4 = $Acknowledged_Table += , @($counter, $r2Array_sum3, $r2Array_status3, $r2Array_Changeby3, $r2Array_URL3)
    $counter ++  
    $A_Table4 = $Acknowledged_Table += , @($counter, $r2Array_sum4, $r2Array_status4, $r2Array_Changeby4, $r2Array_URL4)

    if ($A_Table4) {
        $A_HTML_Table += '<table>'
        $A_HTML_Table += '<tr>'
        $A_HTML_Table += "<th class='Tb_heading' colspan='4'>PagerDuty Acknowledged Alerts</th>"
        $A_HTML_Table += '</<tr>'
        $A_HTML_Table += '<tr>'
        $A_HTML_Table += '<th>Alert Summary</th>'
        $A_HTML_Table += '<th>Status</th>'
        $A_HTML_Table += '<th>Last Changed by</th>'
        $A_HTML_Table += '<th>URL</th>'
        $A_HTML_Table += '</tr>'

        $A_HTML_Table += '<tr>'
        $A_HTML_Table += "<td><p>$r2Array_sum1</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_status1</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_Changeby1</p></td>"
        $A_HTML_Table += "<td><a href='$r2Array_URL1'>$r2Array_URL1</p></td>"
        $A_HTML_Table += '</tr>'

        $A_HTML_Table += '<tr>'
        $A_HTML_Table += "<td><p>$r2Array_sum2</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_status2</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_Changeby2</p></td>"
        $A_HTML_Table += "<td><a href='$r2Array_URL2'>$r2Array_URL2</p></td>"
        $A_HTML_Table += '</tr>'

        $A_HTML_Table += '<tr>'
        $A_HTML_Table += "<td><p>$r2Array_sum3</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_status3</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_Changeby3</p></td>"
        $A_HTML_Table += "<td><a href='$r2Array_URL3'>$r2Array_URL3</p></td>"
        $A_HTML_Table += '</tr>'

        $A_HTML_Table += '<tr>'
        $A_HTML_Table += "<td><p>$r2Array_sum4</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_status4</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_Changeby4</p></td>"
        $A_HTML_Table += "<td><a href='$r2Array_URL4'>$r2Array_URL4</p></td>"
        $A_HTML_Table += '</tr>'
        $A_HTML_Table += '</table>'
    }
    else {
        Write-Host "Not working"
    }
}
else {
    $A_Table5 = $Acknowledged_Table += , @($counter, $r2Array_sum1, $r2Array_status1, $r2Array_Changeby1, $r2Array_URL1)
    $counter ++  
    $A_Table5 = $Acknowledged_Table += , @($counter, $r2Array_sum2, $r2Array_status2, $r2Array_Changeby2, $r2Array_URL2)  
    $counter ++  
    $A_Table5 = $Acknowledged_Table += , @($counter, $r2Array_sum3, $r2Array_status3, $r2Array_Changeby3, $r2Array_URL3)  
    $counter ++  
    $A_Table5 = $Acknowledged_Table += , @($counter, $r2Array_sum4, $r2Array_status4, $r2Array_Changeby4, $r2Array_URL4) 
    $counter ++  
    $A_Table5 = $Acknowledged_Table += , @($counter, $r2Array_sum5, $r2Array_status5, $r2Array_Changeby5, $r2Array_URL5)

    if ($A_Table5) {
        $A_HTML_Table += '<table>'
        $A_HTML_Table += '<tr>'
        $A_HTML_Table += "<th class='Tb_heading' colspan='4'>PagerDuty Acknowledged Alerts</th>"
        $A_HTML_Table += '</<tr>'
        $A_HTML_Table += '<tr>'
        $A_HTML_Table += '<th>Alert Summary</th>'
        $A_HTML_Table += '<th>Status</th>'
        $A_HTML_Table += '<th>Last Changed by</th>'
        $A_HTML_Table += '<th>URL</th>'
        $A_HTML_Table += '</tr>'

        $A_HTML_Table += '<tr>'
        $A_HTML_Table += "<td><p>$r2Array_sum1</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_status1</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_Changeby1</p></td>"
        $A_HTML_Table += "<td><a href='$r2Array_URL1'>$r2Array_URL1</p></td>"
        $A_HTML_Table += '</tr>'

        $A_HTML_Table += '<tr>'
        $A_HTML_Table += "<td><p>$r2Array_sum2</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_status2</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_Changeby2</p></td>"
        $A_HTML_Table += "<td><a href='$r2Array_URL2'>$r2Array_URL2</p></td>"
        $A_HTML_Table += '</tr>'

        $A_HTML_Table += '<tr>'
        $A_HTML_Table += "<td><p>$r2Array_sum3</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_status3</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_Changeby3</p></td>"
        $A_HTML_Table += "<td><a href='$r2Array_URL3'>$r2Array_URL3</p></td>"
        $A_HTML_Table += '</tr>'

        $A_HTML_Table += '<tr>'
        $A_HTML_Table += "<td><p>$r2Array_sum4</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_status4</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_Changeby4</p></td>"
        $A_HTML_Table += "<td><a href='$r2Array_URL4'>$r2Array_URL4</p></td>"
        $A_HTML_Table += '</tr>'

        $A_HTML_Table += '<tr>'
        $A_HTML_Table += "<td><p>$r2Array_sum5</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_status5</p></td>"
        $A_HTML_Table += "<td><p>$r2Array_Changeby5</p></td>"
        $A_HTML_Table += "<td><a href='$r2Array_URL5'>$r2Array_URL5</p></td>"
        $A_HTML_Table += '</tr>'
        $A_HTML_Table += '</table>'
    }
    else {
        Write-Host "Not working"
    }
}
write-host "=============================================================="
Write-Host ""
Write-Host "....................................................................................."
Write-Host ""

foreach ($acknowledged_item in $Acknowledged_Table) {  
    Write-host ($acknowledged_item) # List array content using foreach()
} 
$Acknowledged_Table | % { $_ -join ',' } | Out-File -FilePath ".\Acknowledged_Table.txt"

$HTML_inc_count += "<p>There are <b>$ack_inc_count</b> Acknowledged and <b>$triggered_inc_count</b> Triggered alerts for the NOC.</p>"


if ($triggered_inc_count -eq 0) {
    Clear-Variable T_HTML_Table
}


if ($ack_inc_count -eq 0) {
    Clear-Variable A_HTML_Table
}


Write-Host "Compiling High Severity Table" -ForegroundColor Yellow

# Gets all pen High Severity PD Alerts and creates a new HTML table
$URL_test = "https://api.pagerduty.com/incidents?statuses[]=acknowledged&statuses[]=triggered&limit=200&offset=0"
$res = Invoke-RestMethod -Method Get -Uri $URL_test -Header $headers
$rez = $res.incidents | Select-Object summary, status, priority, html_url



$alert_seva = $rez | Where-Object { $_.priority.summary -match "SevA" }
if ($alert_seva) {
    Write-Host "SevA alerts exist - Compiling into table." -ForegroundColor Green
    $alert_seva | foreach {
        $SevA_Tb += "<tr><td class='SevA_text_left'>$($_.summary)</td> <td class='SevA'>$($_.priority.summary)</td> <td class='SevA'>$($_.status)</td> <td class='SevA'><a href='$($_.html_url)'>$($_.html_url)</td></tr>`r`n"
    }
    foreach ($SevA_alert in $alert_seva) {
        $SevA_alert
    }
}
else {
    write-host "There are no SevA alerts." -ForegroundColor Red
}


$alert_sevb = $rez | Where-Object { $_.priority.summary -match "SevB" }
if ($alert_sevb) {
    Write-Host "SevB alerts exist - Compiling into table." -ForegroundColor Green
    $alert_sevb | foreach {
        $SevB_Tb += "<tr><td class='SevB_text_left'>$($_.summary)</td> <td class='SevB'>$($_.priority.summary)</td> <td class='SevB'>$($_.status)</td> <td class='SevB'><a href='$($_.html_url)'>$($_.html_url)</td></tr>`r`n"
    }
    foreach ($SevB_alert in $alert_sevb) {
        $SevB_alert
    }
}
else {
    write-host "There are no SevB alerts." -ForegroundColor Red
}


$alert_sevc = $rez | Where-Object { $_.priority.summary -match "SevC" }
if ($alert_sevc) {
    Write-Host "SevC alerts exist - Compiling into table." -ForegroundColor Green
    $alert_sevc | foreach {
        $SevC_Tb += "<tr><td class='SevC_text_left'>$($_.summary)</td> <td class='SevC'>$($_.priority.summary)</td> <td class='SevC'>$($_.status)</td> <td class='SevC'><a href='$($_.html_url)'>$($_.html_url)</td></tr>`r`n"
    }
    foreach ($SevC_alert in $alert_sevc) {
        $SevC_alert
    }
}
else {
    write-host "There are no SevC alerts." -ForegroundColor Red
}


$alert_rega = $rez | Where-Object { $_.priority.summary -match "RegA" }
if ($alert_rega) {
    Write-Host "RegA alerts exist - Compiling into table." -ForegroundColor Green
    $alert_rega | foreach {
        $RegA_Tb += "<tr><td class='RegA_text_left'>$($_.summary)</td> <td class='RegA'>$($_.priority.summary)</td> <td class='RegA'>$($_.status)</td> <td class='RegA'><a href='$($_.html_url)'>$($_.html_url)</td></tr>`r`n"
    }
    foreach ($RegA_alert in $alert_rega) {
        $RegA_alert
    }
}
else {
    write-host "There are no RegA alerts." -ForegroundColor Red
}


$alert_regb = $rez | Where-Object { $_.priority.summary -match "RegB" }
if ($alert_regb) {
    Write-Host "RegB alerts exist - Compiling into table." -ForegroundColor Green
    $alert_regb | foreach {
        $RegB_Tb += "<tr><td class='RegB_text_left'>$($_.summary)</td> <td class='RegB'>$($_.priority.summary)</td> <td class='RegB'>$($_.status)</td> <td class='RegB'><a href='$($_.html_url)'>$($_.html_url)</td></tr>`r`n"
    }
    foreach ($RegB_alert in $alert_regb) {
        $RegB_alert
    }
}
else {
    write-host "There are no RegB alerts." -ForegroundColor Red
}


$alert_opa = $rez | Where-Object { $_.priority.summary -match "OpA" }
if ($alert_opa) {
    Write-Host "OpA alerts exist - Compiling into table." -ForegroundColor Green
    $alert_opa | foreach {
        $OpA_Tb += "<tr><td class='OpA_text_left'>$($_.summary)</td> <td class='OpA'>$($_.priority.summary)</td> <td class='OpA'>$($_.status)</td> <td class='OpA'><a href='$($_.html_url)'>$($_.html_url)</td></tr>`r`n"
    }
    foreach ($RegB_alert in $alert_opa) {
        $RegB_alert
    }
}
else {
    write-host "There are no OpA alerts." -ForegroundColor Red
}


if ($alert_seva -or $alert_sevb -or $alert_sevc -or $alert_rega -or $alert_regb -or $alert_opa) {
    $Sev_Tb = "<table>
    <tr>
    <th class='Tb_heading' colspan='4'>Active High Severity Alerts [All Teams]</th>
    </tr>
    <tr>
    <th>Alert Summary</th>
    <th>Severity</th>
    <th>Status</th>
    <th>URL</th>
    </tr>
    $SevA_Tb
    $SevB_Tb
    $SevC_Tb
    $RegA_Tb
    $RegB_Tb
    $OpA_Tb
    </table>"
}
else {
    Clear-Variable Sev_Tb
}

# HTML Construction
$HTMLmessage = @"
<html>
<head>
<style>
    TABLE{border: 1px solid black; border-collapse: collapse; font-size:12pt; font-family: Calibri;}
    TH{border: 1px solid black; background: #333333; padding: 5px; color: #ffffff;}
    TD{border: 1px solid black; padding: 5px; }
    .row{background: #000;}
    a {color: #004de0; text-decoration: none;}
    h1 {font-family: Calibri;}
    p {font-family: Calibri;}
    .null {background-color: #ffffff;}
    .table_heading {font-family: Calibri; padding: 10px; margin: 0;}
    .table_heading_red {font-family: Calibri; padding: 10px; margin: 0; color: red;}
    .Tb_heading {background-color: #333333; padding: 15px; font-size: 18px; letter-spacing: 1px;}
    .bg {background-color: #333333;}
    .SevA_text_left {background-color: rgba(255, 0, 0, 0.205); text-align: left;}
    .SevB_text_left {background-color: #eb61166c; text-align: left;}
    .SevC_text_left {background-color: rgba(255, 220, 42, 0.281); text-align: left;}
    .RegA_text_left {background-color: rgba(104, 26, 177, 0.336); text-align: left;}
    .RegB_text_left {text-align: left;}
    .OpA_text_left {background-color: rgba(0, 76, 142, 0.452); text-align: left;}
    .SevA {background-color: rgba(255, 0, 0, 0.205); text-align: center;}
    .SevB {background-color: #eb61166c; text-align: center;}
    .SevC {background-color: rgba(255, 220, 42, 0.281); text-align: center;}
    .RegA {background-color: rgba(104, 26, 177, 0.336); text-align: center;}
    .RegB {text-align: center;}
    .OpA {background-color: rgba(0, 76, 142, 0.452); text-align: center;}
</style>
<title>Handover Report</title>
</head>
<body>
<h1>Handover Report (PagerDuty Overview)</h1>
"@
$HTMLmessage += $HTML_inc_count
$HTMLmessage += $A_HTML_Table_Heading
$HTMLmessage += $A_HTML_Table
$HTMLmessage += @"
<br>
"@
$HTMLmessage += $T_HTML_Table_Heading
$HTMLmessage += $T_HTML_Table
$HTMLmessage += @"
$Sev_Tb
</body>    
</html>
"@
$HTMLmessage | Out-File -FilePath ".\index.html" # Will Output the results as an index.html file in the same directory
# End of HTML construction and output to index.html file 


# Composing Email Message
$pass = ConvertTo-SecureString "PASSWORD" -AsPlainText -Force
$mycreds = New-Object System.Management.Automation.PSCredential ("USERNAME", $pass)

$From = "EMAIL_ADRESS"
$To = "EMAIL_ADRESS"
$Subject = "PagerDuty Handover Report"
$SMTPServer = "SMTP_SERVER"

try {
    Send-MailMessage -From $From -to $To -Subject $Subject -Body $HTMLmessage -BodyAsHtml -SmtpServer $SMTPServer -Credential $mycreds 
    Write-Host "Email sent!" -ForegroundColor Green
}
catch {
    Write-Host "Failed to send email..." -ForegroundColor Red
}