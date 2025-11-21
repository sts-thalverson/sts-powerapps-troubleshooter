# Automated discovery of flows, plugins, and processes related to project approvals
# Focuses on finding email automation that stopped working around July 8th

param(
    [string]$EnvironmentUrl = "https://dhapd.crm11.dynamics.com",
    [string]$SearchKeywords = "project,approval,approve,email,notification",
    [datetime]$BreakageDate = (Get-Date "2024-07-08")
)

$Keywords = $SearchKeywords -split ',' | ForEach-Object { $_.Trim().ToLower() }

# Install required modules
if (!(Get-Module -ListAvailable -Name Microsoft.Xrm.Tooling.CrmConnector.PowerShell)) {
    Write-Host "Installing Microsoft.Xrm.Tooling.CrmConnector.PowerShell module..." -ForegroundColor Yellow
    Install-Module -Name Microsoft.Xrm.Tooling.CrmConnector.PowerShell -Scope CurrentUser -Force -AllowClobber
}

if (!(Get-Module -ListAvailable -Name Microsoft.Xrm.Data.PowerShell)) {
    Write-Host "Installing Microsoft.Xrm.Data.PowerShell module..." -ForegroundColor Yellow
    Install-Module -Name Microsoft.Xrm.Data.PowerShell -Scope CurrentUser -Force -AllowClobber
}

Import-Module Microsoft.Xrm.Tooling.CrmConnector.PowerShell
Import-Module Microsoft.Xrm.Data.PowerShell

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "AUTOMATED FLOW DISCOVERY TOOL" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Searching for: $SearchKeywords" -ForegroundColor Yellow
Write-Host "Breakage date: $($BreakageDate.ToString('yyyy-MM-dd'))" -ForegroundColor Yellow
Write-Host ""

# Connect with MFA support
if (-not $EnvironmentUrl.StartsWith("http")) {
    $EnvironmentUrl = "https://$EnvironmentUrl"
}

$connectionString = "AuthType=OAuth;Url=$EnvironmentUrl;AppId=51f81489-12ee-4a9e-aaae-a2591f45987d;RedirectUri=http://localhost;LoginPrompt=Auto"
$conn = Get-CrmConnection -ConnectionString $connectionString

if (-not $conn.IsReady) {
    Write-Error "Failed to connect to Dataverse"
    exit
}

Write-Host "Connected to: $($conn.ConnectedOrgFriendlyName)" -ForegroundColor Green
Write-Host ""

# Results collection
$discoveredItems = @()

# ============================================================================
# 1. SCAN ALL CLOUD FLOWS (Power Automate)
# ============================================================================
Write-Host "=== SCANNING CLOUD FLOWS ===" -ForegroundColor Cyan

$allFlowsFetch = @"
<fetch>
  <entity name='workflow'>
    <attribute name='name' />
    <attribute name='workflowid' />
    <attribute name='description' />
    <attribute name='statecode' />
    <attribute name='statuscode' />
    <attribute name='category' />
    <attribute name='primaryentity' />
    <attribute name='createdon' />
    <attribute name='modifiedon' />
    <attribute name='clientdata' />
    <filter type='and'>
      <condition attribute='category' operator='eq' value='5' />
      <condition attribute='type' operator='eq' value='1' />
    </filter>
    <order attribute='modifiedon' descending='true' />
  </entity>
</fetch>
"@

$allFlows = Get-CrmRecordsByFetch -conn $conn -Fetch $allFlowsFetch
Write-Host "Found $($allFlows.CrmRecords.Count) total cloud flows" -ForegroundColor Gray

foreach ($flow in $allFlows.CrmRecords) {
    $matchScore = 0
    $matchReasons = @()
    
    # Check name for keywords
    $flowName = $flow.name.ToLower()
    foreach ($keyword in $Keywords) {
        if ($flowName -like "*$keyword*") {
            $matchScore += 2
            $matchReasons += "Name contains '$keyword'"
        }
    }
    
    # Check description
    if ($flow.description) {
        $flowDesc = $flow.description.ToLower()
        foreach ($keyword in $Keywords) {
            if ($flowDesc -like "*$keyword*") {
                $matchScore += 1
                $matchReasons += "Description contains '$keyword'"
            }
        }
    }
    
    # Check clientdata (flow definition JSON)
    if ($flow.clientdata) {
        $clientData = $flow.clientdata.ToLower()
        
        # Check for email-related actions
        if ($clientData -match 'send.*email|office365|smtp|mail\.send') {
            $matchScore += 3
            $matchReasons += "Contains email sending action"
        }
        
        # Check for project entity references
        if ($clientData -match 'msdyn_project|project') {
            $matchScore += 2
            $matchReasons += "References project entity"
        }
        
        # Check for approval actions
        if ($clientData -match 'approval|approve|reject') {
            $matchScore += 2
            $matchReasons += "Contains approval logic"
        }
        
        # Check for keywords in flow definition
        foreach ($keyword in $Keywords) {
            if ($clientData -like "*$keyword*") {
                $matchScore += 1
            }
        }
    }
    
    # Check if modified around breakage date (within 30 days)
    if ($flow.modifiedon) {
        try {
            $modDate = [DateTime]::Parse($flow.modifiedon)
            $daysSinceBreakage = ($modDate - $BreakageDate).Days
            if ([Math]::Abs($daysSinceBreakage) -le 30) {
                $matchScore += 2
                $matchReasons += "Modified around breakage date ($daysSinceBreakage days)"
            }
        } catch {
            # Skip date comparison if parsing fails
        }
    }
    
    if ($matchScore -gt 0) {
        $state = if ($flow.statecode -eq 1) { "On" } else { "Off" }
        
        $discoveredItems += [PSCustomObject]@{
            Type = "Cloud Flow"
            Name = $flow.name
            Id = $flow.workflowid
            State = $state
            ModifiedOn = $flow.modifiedon
            Score = $matchScore
            Reasons = ($matchReasons -join "; ")
            PrimaryEntity = $flow.primaryentity
        }
    }
}

# ============================================================================
# 2. SCAN CLASSIC WORKFLOWS
# ============================================================================
Write-Host "=== SCANNING CLASSIC WORKFLOWS ===" -ForegroundColor Cyan

$classicWorkflowsFetch = @"
<fetch>
  <entity name='workflow'>
    <attribute name='name' />
    <attribute name='workflowid' />
    <attribute name='description' />
    <attribute name='statecode' />
    <attribute name='xaml' />
    <attribute name='modifiedon' />
    <filter type='and'>
      <condition attribute='type' operator='eq' value='1' />
      <condition attribute='category' operator='ne' value='5' />
      <condition attribute='category' operator='ne' value='4' />
    </filter>
    <order attribute='modifiedon' descending='true' />
  </entity>
</fetch>
"@

$classicWorkflows = Get-CrmRecordsByFetch -conn $conn -Fetch $classicWorkflowsFetch
Write-Host "Found $($classicWorkflows.CrmRecords.Count) classic workflows" -ForegroundColor Gray

foreach ($wf in $classicWorkflows.CrmRecords) {
    $matchScore = 0
    $matchReasons = @()
    
    $wfName = $wf.name.ToLower()
    foreach ($keyword in $Keywords) {
        if ($wfName -like "*$keyword*") {
            $matchScore += 2
            $matchReasons += "Name contains '$keyword'"
        }
    }
    
    # Check XAML for email activities
    if ($wf.xaml) {
        $xaml = $wf.xaml.ToLower()
        
        if ($xaml -match 'sendemail|createemail') {
            $matchScore += 3
            $matchReasons += "Contains email action"
        }
        
        if ($xaml -match 'msdyn_project') {
            $matchScore += 2
            $matchReasons += "Works with project entity"
        }
        
        foreach ($keyword in $Keywords) {
            if ($xaml -like "*$keyword*") {
                $matchScore += 1
            }
        }
    }
    
    if ($wf.modifiedon) {
        try {
            $modDate = [DateTime]::Parse($wf.modifiedon)
            $daysSinceBreakage = ($modDate - $BreakageDate).Days
            if ([Math]::Abs($daysSinceBreakage) -le 30) {
                $matchScore += 2
                $matchReasons += "Modified around breakage date ($daysSinceBreakage days)"
            }
        } catch {
            # Skip date comparison if parsing fails
        }
    }
    
    if ($matchScore -gt 0) {
        $state = if ($wf.statecode -eq 1) { "On" } else { "Off" }
        
        $discoveredItems += [PSCustomObject]@{
            Type = "Classic Workflow"
            Name = $wf.name
            Id = $wf.workflowid
            State = $state
            ModifiedOn = $wf.modifiedon
            Score = $matchScore
            Reasons = ($matchReasons -join "; ")
            PrimaryEntity = ""
        }
    }
}

# ============================================================================
# 3. SCAN PLUGINS
# ============================================================================
Write-Host "=== SCANNING PLUGINS ===" -ForegroundColor Cyan

$pluginsFetch = @"
<fetch>
  <entity name='plugintype'>
    <attribute name='typename' />
    <attribute name='plugintypeid' />
    <attribute name='friendlyname' />
    <link-entity name='sdkmessageprocessingstep' from='plugintypeid' to='plugintypeid' alias='step'>
      <attribute name='name' />
      <attribute name='sdkmessageprocessingstepid' />
      <attribute name='statecode' />
      <attribute name='stage' />
      <attribute name='mode' />
      <attribute name='rank' />
      <link-entity name='sdkmessagefilter' from='sdkmessagefilterid' to='sdkmessagefilterid' alias='filter'>
        <attribute name='primaryobjecttypecode' />
      </link-entity>
      <link-entity name='sdkmessage' from='sdkmessageid' to='sdkmessageid' alias='message'>
        <attribute name='name' />
      </link-entity>
    </link-entity>
  </entity>
</fetch>
"@

try {
    $plugins = Get-CrmRecordsByFetch -conn $conn -Fetch $pluginsFetch
    Write-Host "Found $($plugins.CrmRecords.Count) plugin steps" -ForegroundColor Gray
    
    foreach ($plugin in $plugins.CrmRecords) {
        $matchScore = 0
        $matchReasons = @()
        
        # Check if it's on project entity
        $entityCode = $plugin.'filter.primaryobjecttypecode'
        if ($entityCode -eq 10054) {  # msdyn_project object type code
            $matchScore += 3
            $matchReasons += "Registered on Project entity"
        }
        
        # Check step name
        if ($plugin.'step.name') {
            $stepName = $plugin.'step.name'.ToLower()
            foreach ($keyword in $Keywords) {
                if ($stepName -like "*$keyword*") {
                    $matchScore += 2
                    $matchReasons += "Step name contains '$keyword'"
                }
            }
        }
        
        # Check message type
        $messageName = $plugin.'message.name'
        if ($messageName -in @('Create', 'Update', 'Delete')) {
            if ($matchScore -gt 0) {
                $matchReasons += "Triggers on $messageName"
            }
        }
        
        if ($matchScore -gt 0) {
            $state = if ($plugin.'step.statecode' -eq 0) { "On" } else { "Off" }
            $stage = switch ($plugin.'step.stage') {
                10 { "Pre-validation" }
                20 { "Pre-operation" }
                40 { "Post-operation" }
                default { "Unknown" }
            }
            
            $discoveredItems += [PSCustomObject]@{
                Type = "Plugin"
                Name = "$($plugin.typename) - $($plugin.'step.name')"
                Id = $plugin.'step.sdkmessageprocessingstepid'
                State = $state
                ModifiedOn = $null
                Score = $matchScore
                Reasons = ($matchReasons -join "; ") + " | Stage: $stage | Message: $messageName"
                PrimaryEntity = $entityCode
            }
        }
    }
} catch {
    Write-Host "Error scanning plugins: $_" -ForegroundColor Red
}

# ============================================================================
# 4. SCAN WEB RESOURCES (JavaScript/HTML that might call flows or send emails)
# ============================================================================
Write-Host "=== SCANNING WEB RESOURCES ===" -ForegroundColor Cyan

$webResourcesFetch = @"
<fetch>
  <entity name='webresource'>
    <attribute name='name' />
    <attribute name='webresourceid' />
    <attribute name='displayname' />
    <attribute name='webresourcetype' />
    <attribute name='content' />
    <attribute name='modifiedon' />
    <filter type='and'>
      <filter type='or'>
        <condition attribute='webresourcetype' operator='eq' value='3' />
        <condition attribute='webresourcetype' operator='eq' value='1' />
        <condition attribute='webresourcetype' operator='eq' value='2' />
      </filter>
    </filter>
    <order attribute='modifiedon' descending='true' />
  </entity>
</fetch>
"@

try {
    $webResources = Get-CrmRecordsByFetch -conn $conn -Fetch $webResourcesFetch
    Write-Host "Found $($webResources.CrmRecords.Count) web resources (JS/HTML/CSS)" -ForegroundColor Gray
    
    foreach ($wr in $webResources.CrmRecords) {
        $matchScore = 0
        $matchReasons = @()
        
        $wrName = $wr.name.ToLower()
        $wrDisplayName = if ($wr.displayname) { $wr.displayname.ToLower() } else { "" }
        
        # Check name and display name
        foreach ($keyword in $Keywords) {
            if ($wrName -like "*$keyword*" -or $wrDisplayName -like "*$keyword*") {
                $matchScore += 2
                $matchReasons += "Name contains '$keyword'"
            }
        }
        
        # Decode and check content (web resources are base64 encoded)
        if ($wr.content) {
            try {
                $decodedContent = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($wr.content)).ToLower()
                
                # Check for flow/API calls
                if ($decodedContent -match 'fetch.*api.*dynamics|webapi|executeaction|executeworkflow') {
                    $matchScore += 3
                    $matchReasons += "Makes API/workflow calls"
                }
                
                # Check for email-related code
                if ($decodedContent -match 'email|smtp|sendmail|mailto') {
                    $matchScore += 2
                    $matchReasons += "Contains email-related code"
                }
                
                # Check for project entity references
                if ($decodedContent -match 'msdyn_project|projects?') {
                    $matchScore += 2
                    $matchReasons += "References project entity"
                }
                
                # Check for keywords
                foreach ($keyword in $Keywords) {
                    if ($decodedContent -like "*$keyword*") {
                        $matchScore += 1
                    }
                }
            } catch {
                # Skip if content can't be decoded
            }
        }
        
        # Check modification date
        if ($wr.modifiedon) {
            try {
                $modDate = [DateTime]::Parse($wr.modifiedon)
                $daysSinceBreakage = ($modDate - $BreakageDate).Days
                if ([Math]::Abs($daysSinceBreakage) -le 30) {
                    $matchScore += 2
                    $matchReasons += "Modified around breakage date ($daysSinceBreakage days)"
                }
            } catch {
                # Skip date comparison if parsing fails
            }
        }
        
        if ($matchScore -gt 0) {
            $wrType = switch ($wr.webresourcetype_Property.Value.Value) {
                1 { "HTML" }
                2 { "CSS" }
                3 { "JavaScript" }
                4 { "XML" }
                5 { "PNG" }
                6 { "JPG" }
                7 { "GIF" }
                8 { "XAP" }
                9 { "XSL" }
                10 { "ICO" }
                default { "Unknown" }
            }
            
            $discoveredItems += [PSCustomObject]@{
                Type = "Web Resource ($wrType)"
                Name = $wr.displayname
                Id = $wr.webresourceid
                State = "N/A"
                ModifiedOn = $wr.modifiedon
                Score = $matchScore
                Reasons = ($matchReasons -join "; ")
                PrimaryEntity = ""
            }
        }
    }
} catch {
    Write-Host "Error scanning web resources: $_" -ForegroundColor Red
}

Write-Host ""

# ============================================================================
# 5. SCAN CUSTOM ACTIONS
# ============================================================================
Write-Host "=== SCANNING CUSTOM ACTIONS ===" -ForegroundColor Cyan

$customActionsFetch = @"
<fetch>
  <entity name='workflow'>
    <attribute name='name' />
    <attribute name='workflowid' />
    <attribute name='description' />
    <attribute name='statecode' />
    <attribute name='xaml' />
    <attribute name='modifiedon' />
    <filter type='and'>
      <condition attribute='type' operator='eq' value='1' />
      <condition attribute='category' operator='eq' value='3' />
    </filter>
    <order attribute='modifiedon' descending='true' />
  </entity>
</fetch>
"@

try {
    $customActions = Get-CrmRecordsByFetch -conn $conn -Fetch $customActionsFetch
    Write-Host "Found $($customActions.CrmRecords.Count) custom actions" -ForegroundColor Gray
    
    foreach ($action in $customActions.CrmRecords) {
        $matchScore = 0
        $matchReasons = @()
        
        $actionName = $action.name.ToLower()
        foreach ($keyword in $Keywords) {
            if ($actionName -like "*$keyword*") {
                $matchScore += 2
                $matchReasons += "Name contains '$keyword'"
            }
        }
        
        # Check XAML for email activities
        if ($action.xaml) {
            $xaml = $action.xaml.ToLower()
            
            if ($xaml -match 'sendemail|createemail') {
                $matchScore += 3
                $matchReasons += "Contains email action"
            }
            
            if ($xaml -match 'msdyn_project') {
                $matchScore += 2
                $matchReasons += "Works with project entity"
            }
            
            foreach ($keyword in $Keywords) {
                if ($xaml -like "*$keyword*") {
                    $matchScore += 1
                }
            }
        }
        
        if ($action.modifiedon) {
            try {
                $modDate = [DateTime]::Parse($action.modifiedon)
                $daysSinceBreakage = ($modDate - $BreakageDate).Days
                if ([Math]::Abs($daysSinceBreakage) -le 30) {
                    $matchScore += 2
                    $matchReasons += "Modified around breakage date ($daysSinceBreakage days)"
                }
            } catch {
                # Skip date comparison if parsing fails
            }
        }
        
        if ($matchScore -gt 0) {
            $state = if ($action.statecode -eq 1) { "On" } else { "Off" }
            
            $discoveredItems += [PSCustomObject]@{
                Type = "Custom Action"
                Name = $action.name
                Id = $action.workflowid
                State = $state
                ModifiedOn = $action.modifiedon
                Score = $matchScore
                Reasons = ($matchReasons -join "; ")
                PrimaryEntity = ""
            }
        }
    }
} catch {
    Write-Host "Error scanning custom actions: $_" -ForegroundColor Red
}

Write-Host ""

# ============================================================================
# 6. SCAN SERVICE ENDPOINTS (Webhooks/Azure Service Bus)
# ============================================================================
Write-Host "=== SCANNING SERVICE ENDPOINTS ===" -ForegroundColor Cyan

$serviceEndpointsFetch = @"
<fetch>
  <entity name='serviceendpoint'>
    <attribute name='name' />
    <attribute name='serviceendpointid' />
    <attribute name='description' />
    <attribute name='url' />
    <attribute name='contract' />
    <attribute name='modifiedon' />
    <link-entity name='sdkmessageprocessingstep' from='eventhandler' to='serviceendpointid' alias='step'>
      <attribute name='name' />
      <attribute name='sdkmessageprocessingstepid' />
      <attribute name='statecode' />
      <link-entity name='sdkmessagefilter' from='sdkmessagefilterid' to='sdkmessagefilterid' alias='filter'>
        <attribute name='primaryobjecttypecode' />
      </link-entity>
    </link-entity>
  </entity>
</fetch>
"@

try {
    $serviceEndpoints = Get-CrmRecordsByFetch -conn $conn -Fetch $serviceEndpointsFetch
    Write-Host "Found $($serviceEndpoints.CrmRecords.Count) service endpoint registrations" -ForegroundColor Gray
    
    foreach ($endpoint in $serviceEndpoints.CrmRecords) {
        $matchScore = 0
        $matchReasons = @()
        
        # Check if it's on project entity
        $entityCode = $endpoint.'filter.primaryobjecttypecode'
        if ($entityCode -eq 10054) {  # msdyn_project
            $matchScore += 3
            $matchReasons += "Registered on Project entity"
        }
        
        $endpointName = $endpoint.name.ToLower()
        foreach ($keyword in $Keywords) {
            if ($endpointName -like "*$keyword*") {
                $matchScore += 2
                $matchReasons += "Name contains '$keyword'"
            }
        }
        
        # Check URL for external services
        if ($endpoint.url) {
            $url = $endpoint.url.ToLower()
            if ($url -match 'logic.*app|function.*app|servicebus|webhook') {
                $matchScore += 2
                $matchReasons += "Connects to external automation service"
            }
        }
        
        if ($endpoint.modifiedon) {
            try {
                $modDate = [DateTime]::Parse($endpoint.modifiedon)
                $daysSinceBreakage = ($modDate - $BreakageDate).Days
                if ([Math]::Abs($daysSinceBreakage) -le 30) {
                    $matchScore += 2
                    $matchReasons += "Modified around breakage date ($daysSinceBreakage days)"
                }
            } catch {
                # Skip date comparison if parsing fails
            }
        }
        
        if ($matchScore -gt 0) {
            $state = if ($endpoint.'step.statecode' -eq 0) { "On" } else { "Off" }
            $contract = switch ($endpoint.contract) {
                1 { "OneWay" }
                2 { "Queue" }
                3 { "Rest" }
                4 { "TwoWay" }
                5 { "Topic" }
                6 { "Webhook" }
                7 { "EventHub" }
                8 { "EventGrid" }
                default { "Unknown" }
            }
            
            $discoveredItems += [PSCustomObject]@{
                Type = "Service Endpoint ($contract)"
                Name = "$($endpoint.name) - $($endpoint.'step.name')"
                Id = $endpoint.'step.sdkmessageprocessingstepid'
                State = $state
                ModifiedOn = $endpoint.modifiedon
                Score = $matchScore
                Reasons = ($matchReasons -join "; ") + " | URL: $($endpoint.url)"
                PrimaryEntity = $entityCode
            }
        }
    }
} catch {
    Write-Host "Error scanning service endpoints: $_" -ForegroundColor Red
}

Write-Host ""

# ============================================================================
# 7. SCAN VIRTUAL ENTITIES (Data Sources that might trigger external automation)
# ============================================================================
Write-Host "=== SCANNING VIRTUAL ENTITY DATA SOURCES ===" -ForegroundColor Cyan

$dataSourcesFetch = @"
<fetch>
  <entity name='entitydatasource'>
    <attribute name='name' />
    <attribute name='entitydatasourceid' />
    <attribute name='description' />
    <attribute name='modifiedon' />
  </entity>
</fetch>
"@

try {
    $dataSources = Get-CrmRecordsByFetch -conn $conn -Fetch $dataSourcesFetch
    Write-Host "Found $($dataSources.CrmRecords.Count) virtual entity data sources" -ForegroundColor Gray
    
    foreach ($ds in $dataSources.CrmRecords) {
        $matchScore = 0
        $matchReasons = @()
        
        $dsName = $ds.name.ToLower()
        foreach ($keyword in $Keywords) {
            if ($dsName -like "*$keyword*") {
                $matchScore += 2
                $matchReasons += "Name contains '$keyword'"
            }
        }
        
        if ($matchScore -gt 0) {
            $discoveredItems += [PSCustomObject]@{
                Type = "Virtual Entity Data Source"
                Name = $ds.name
                Id = $ds.entitydatasourceid
                State = "N/A"
                ModifiedOn = $ds.modifiedon
                Score = $matchScore
                Reasons = ($matchReasons -join "; ")
                PrimaryEntity = ""
            }
        }
    }
} catch {
    Write-Host "Error scanning data sources: $_" -ForegroundColor Red
}

Write-Host ""

# ============================================================================
# 8. SCAN SLAs (Service Level Agreements - can trigger actions)
# ============================================================================
Write-Host "=== SCANNING SLAs ===" -ForegroundColor Cyan

$slasFetch = @"
<fetch>
  <entity name='sla'>
    <attribute name='name' />
    <attribute name='slaid' />
    <attribute name='description' />
    <attribute name='statecode' />
    <attribute name='modifiedon' />
    <attribute name='objecttypecode' />
  </entity>
</fetch>
"@

try {
    $slas = Get-CrmRecordsByFetch -conn $conn -Fetch $slasFetch
    Write-Host "Found $($slas.CrmRecords.Count) SLAs" -ForegroundColor Gray
    
    foreach ($sla in $slas.CrmRecords) {
        $matchScore = 0
        $matchReasons = @()
        
        # Check if on project entity
        if ($sla.objecttypecode -eq 10054) {
            $matchScore += 2
            $matchReasons += "Applied to Project entity"
        }
        
        $slaName = $sla.name.ToLower()
        foreach ($keyword in $Keywords) {
            if ($slaName -like "*$keyword*") {
                $matchScore += 2
                $matchReasons += "Name contains '$keyword'"
            }
        }
        
        if ($matchScore -gt 0) {
            $state = if ($sla.statecode -eq 0) { "Active" } else { "Inactive" }
            
            $discoveredItems += [PSCustomObject]@{
                Type = "SLA"
                Name = $sla.name
                Id = $sla.slaid
                State = $state
                ModifiedOn = $sla.modifiedon
                Score = $matchScore
                Reasons = ($matchReasons -join "; ")
                PrimaryEntity = $sla.objecttypecode
            }
        }
    }
} catch {
    Write-Host "Error scanning SLAs: $_" -ForegroundColor Red
}

Write-Host ""

# ============================================================================
# 9. SCAN BUSINESS PROCESS FLOWS
# ============================================================================
Write-Host "=== SCANNING BUSINESS PROCESS FLOWS ===" -ForegroundColor Cyan

$bpfsFetch = @"
<fetch>
  <entity name='workflow'>
    <attribute name='name' />
    <attribute name='workflowid' />
    <attribute name='statecode' />
    <attribute name='modifiedon' />
    <filter type='and'>
      <condition attribute='category' operator='eq' value='4' />
    </filter>
  </entity>
</fetch>
"@

$bpfs = Get-CrmRecordsByFetch -conn $conn -Fetch $bpfsFetch
Write-Host "Found $($bpfs.CrmRecords.Count) business process flows" -ForegroundColor Gray

foreach ($bpf in $bpfs.CrmRecords) {
    $matchScore = 0
    $matchReasons = @()
    
    $bpfName = $bpf.name.ToLower()
    foreach ($keyword in $Keywords) {
        if ($bpfName -like "*$keyword*") {
            $matchScore += 2
            $matchReasons += "Name contains '$keyword'"
        }
    }
    
    if ($matchScore -gt 0) {
        $state = if ($bpf.statecode -eq 1) { "On" } else { "Off" }
        
        $discoveredItems += [PSCustomObject]@{
            Type = "Business Process Flow"
            Name = $bpf.name
            Id = $bpf.workflowid
            State = $state
            ModifiedOn = $bpf.modifiedon
            Score = $matchScore
            Reasons = ($matchReasons -join "; ")
            PrimaryEntity = ""
        }
    }
}

# ============================================================================
# 5. ANALYZE AND DISPLAY RESULTS
# ============================================================================
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "DISCOVERY RESULTS" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

if ($discoveredItems.Count -eq 0) {
    Write-Host "No matching items found" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "This suggests the automation may be:" -ForegroundColor Yellow
    Write-Host "  1. Using environment variables or connection references for the sender email" -ForegroundColor Gray
    Write-Host "  2. Configured in a child flow that's called by a parent flow" -ForegroundColor Gray
    Write-Host "  3. Using a custom connector or API" -ForegroundColor Gray
    Write-Host "  4. Handled outside of Dataverse (e.g., Logic Apps, Azure Functions)" -ForegroundColor Gray
} else {
    # Sort by score descending
    $sorted = $discoveredItems | Sort-Object -Property Score -Descending
    
    Write-Host "Found $($discoveredItems.Count) potentially relevant items" -ForegroundColor Green
    Write-Host ""
    
    # Display top candidates
    Write-Host "TOP CANDIDATES (sorted by relevance score):" -ForegroundColor Yellow
    Write-Host ""
    
    foreach ($item in $sorted | Select-Object -First 10) {
        $color = switch ($item.State) {
            "On" { "Green" }
            "Off" { "Red" }
            default { "Yellow" }
        }
        
        Write-Host "[$($item.Score) points] " -NoNewline -ForegroundColor Cyan
        Write-Host "$($item.Type): " -NoNewline -ForegroundColor White
        Write-Host "$($item.Name)" -ForegroundColor $color
        Write-Host "  State: $($item.State)" -ForegroundColor $color
        Write-Host "  ID: $($item.Id)" -ForegroundColor Gray
        if ($item.ModifiedOn) {
            Write-Host "  Modified: $($item.ModifiedOn)" -ForegroundColor Gray
        }
        Write-Host "  Match reasons: $($item.Reasons)" -ForegroundColor Gray
        Write-Host ""
    }
    
    # Export to CSV
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $exportPath = ".\FlowDiscovery_$timestamp.csv"
    $sorted | Export-Csv -Path $exportPath -NoTypeInformation
    Write-Host "Full results exported to: $exportPath" -ForegroundColor Cyan
    
    # Generate Power Automate URLs for flows
    Write-Host "`n=== DIRECT LINKS TO INVESTIGATE ===" -ForegroundColor Cyan
    $envId = $conn.CrmConnectOrgUriActual.Host -replace '\.crm.*\.dynamics\.com', ''
    
    foreach ($item in $sorted | Where-Object { $_.Type -eq "Cloud Flow" } | Select-Object -First 5) {
        Write-Host "Flow: $($item.Name)" -ForegroundColor White
        Write-Host "  https://make.powerautomate.com/environments/$envId/flows/$($item.Id)/details" -ForegroundColor Cyan
        Write-Host ""
    }
}

Write-Host "`n=== RECOMMENDED ACTIONS ===" -ForegroundColor Cyan
Write-Host "1. Review the top candidates listed above" -ForegroundColor White
Write-Host "2. Check flow run history around $($BreakageDate.ToString('yyyy-MM-dd'))" -ForegroundColor White
Write-Host "3. Look for failed runs or connection errors" -ForegroundColor White
Write-Host "4. Verify email connector settings and connection references" -ForegroundColor White
Write-Host "5. Check if any flows were turned off or modified recently" -ForegroundColor White
Write-Host ""