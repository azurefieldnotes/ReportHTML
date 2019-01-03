<#PSScriptInfo

.VERSION 1.0

.GUID 89267615-ea56-44a8-8465-6203b930e4df

.AUTHOR matt.quickenden

.COMPANYNAME ACE

.COPYRIGHT 

.TAGS 

.LICENSEURI 

.PROJECTURI http://www.azurefieldnotes.com/2018/02/08/reporting-on-resource-group-tags-in-azure

.ICONURI 

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES


#>

<# 

.DESCRIPTION 
A powerful module for creating HTML reports within PowerShell no HTML coding required.  

For more details on what is possible, you can view the help file 
https://azurefieldnotesblog.blob.core.windows.net/wp-content/2017/06/Help-ReportHTML2.html 

There is a multi-part blog series.  (which is a little out of date but still relevant) http://www.azurefieldnotes.com/2016/08/04/powershellhtmlreportingpart1

#> 

#####Requires –Modules ReportHTML

Param
(
    [parameter(Mandatory=$false,ValueFromPipeline = $true)] 
    [Array]$KeyNames = @('Owner','Solution'),
    [parameter(Mandatory=$false)] 
    [String]$ReportOutputPath=$env:temp,
    [parameter(Mandatory=$false)] 
    [String]$YouLogoHereURLString , 
    [parameter(Mandatory=$false)] 
    [string]$SubscriptionID,
    $skipSize = 1
    #$EmailDomain = "*"
)

#Login-AzureRmAccount

function Get-AzureRmCachedAccessToken()
{
  $ErrorActionPreference = 'Stop'
  
  if(-not (Get-Module AzureRm.Profile)) {
    Import-Module AzureRm.Profile
  }
  $azureRmProfileModuleVersion = (Get-Module AzureRm.Profile).Version
  # refactoring performed in AzureRm.Profile v3.0 or later
  if($azureRmProfileModuleVersion.Major -ge 3) {
    $azureRmProfile = [Microsoft.Azure.Commands.Common.Authentication.Abstractions.AzureRmProfileProvider]::Instance.Profile
    if(-not $azureRmProfile.Accounts.Count) {
      Write-Error "Ensure you have logged in before calling this function."    
    }
  } else {
    # AzureRm.Profile < v3.0
    $azureRmProfile = [Microsoft.WindowsAzure.Commands.Common.AzureRmProfileProvider]::Instance.Profile
    if(-not $azureRmProfile.Context.Account.Count) {
      Write-Error "Ensure you have logged in before calling this function."    
    }
  }
  
  $currentAzureContext = Get-AzureRmContext
  $profileClient = New-Object Microsoft.Azure.Commands.ResourceManager.Common.RMProfileClient($azureRmProfile)
  Write-Debug ("Getting access token for tenant" + $currentAzureContext.Subscription.TenantId)
  $token = $profileClient.AcquireAccessToken($currentAzureContext.Subscription.TenantId)
  $token.AccessToken
}

#[switch]$AutoKeyName =$false
#$m = Get-Module -List ReportHTML
#if(!$m) {"Can't locate module ReportHTML.  Use Install-module ReportHTML";break}
#else {import-module reporthtml}

if ([string]::IsNullOrEmpty($(Get-AzureRmContext).Account)) {Login-AzureRmAccount}

if ($SubscriptionID -eq '') {
    $RGs = @()
    $subs = @(Get-AzureRmSubscription)
    foreach ($sub in $subs)
    {
        Write-Verbose "selecting subscription $($sub.SubscriptionId) $($sub.Name)"
        Select-AzureRmSubscription $sub.SubscriptionId
        $RGs += Get-AzureRmResourceGroup    
    }

}
else
{
    Select-AzureRmSubscription $SubscriptionID
    $RGs = Get-AzureRmResourceGroup 
}

if ($KeyNames.count -eq 0) 
{
    [switch]$AutoKeyName =$true
    $KeyNames = (($rgs.Tags.keys) | select -Unique)
}

$SubscriptionRGs = @()
foreach ($RG in $RGs) 
{

    $myRG = [PSCustomObject]@{
        ResourceGroupId = $RG.ResourceId
        SubscriptionId = $RG.ResourceId.Split('/')[2]
        Subscription = ($SUBS | ? {$_.id -eq $RG.ResourceId.Split('/')[2]}).name
        ResourceGroupName     = $RG.ResourceGroupName
        Location = $RG.Location
        Link    =  ("URL01" + "https://portal.azure.com/#resource" + $RG.ResourceId  +  "URL02" + ($RG.ResourceId.Split('/') | select -last 1) + "URL03" )
    }

    $i=0
    foreach ($KeyName in $KeyNames) 
    {
        if ($AutoKeyName)
        {
            $myRG | Add-Member -MemberType NoteProperty -Name ([string]$i + "_" + $keyname) -Value $rg.Tags.($KeyName)
            $i++
        }
        else
        {
            $myRG | Add-Member -MemberType NoteProperty -Name ($keyname) -Value $rg.Tags.($KeyName)
        }
    }
    $SubscriptionRGs += $myRG
}

$year =(get-date).year
$month =(get-date).Month
$DaysInMonth= [DateTime]::DaysInMonth($year, $month )

$token =  Get-AzureRmCachedAccessToken
$headers = @{"authorization"="bearer $token"}
$Body =  @{"type"="Usage";"timeframe"="Custom";"timePeriod"=@{"from"="$($year)-$($month)-01T00:00:00+00:00";"to"="$($year)-$($month)-$($DaysInMonth)T23:59:59+00:00"};"dataSet"=@{"granularity"="Daily";"aggregation"=@{"totalCost"=@{"name"="PreTaxCost";"function"="Sum"}};"sorting"=@(@{"direction"="ascending";"name"="UsageDate"})}} 

$Records = @()
#$SubscriptionRG = $SubscriptionRGs[5]
foreach ($SubscriptionRG in  $SubscriptionRGs ) 
{
    $usageUri = "https://management.azure.com/subscriptions/$($SubscriptionRG.SubscriptionId)/resourceGroups/$($SubscriptionRG.ResourceGroupName)/providers/Microsoft.CostManagement/query?api-version=2018-08-31"   
    $Record = '' | select usage, SubscriptionRG
    $Record.usage = Invoke-RestMethod $usageUri -Headers $headers -ContentType "application/json" -Method Post -Body ($body | ConvertTo-Json -Depth 100) 
    $Record.SubscriptionRG = $SubscriptionRG
    $Records += $Record
}

$DataSetDateRange = $Records.usage.properties.rows | % {$_[1]} | select -Unique
$RGCosts =@()
#$usageResult = $usageResults[10]
foreach ($Record in $Records)
{ 
    $RGCost = '' | select SubscriptionID, Subscription,Location, Link, ResourceGroupName, Owner,OwnerName, Solution,OwnerSolution, UsageData ,TotalCost
    $UsageData = @()
    foreach ($DataDate in $DataSetDateRange)
    {
        
        $NewDate = ([string]$DataDate ).substring(6,2) +'/' + ([string]$DataDate ).substring(4,2)  +'/' + ([string]$DataDate ).substring(2,2)
        $AllRows = $Record.usage.properties.rows 
        $Currency = $Record.usage.properties.rows[0][2]
        $FoundRecord = $AllRows | ? {$_[1] -eq $DataDate}
        $Usage = '' | select Cost, Date, Currency
        if ($FoundRecord.Count -eq 3) 
        {
            $Usage.Cost = [Math]::round(([float]$FoundRecord[0]),1)
            $Usage.Date = $NewDate
            $Usage.Currency = $Currency
        }
        else
        {
            $Usage.Cost = 0
            $Usage.Date = $NewDate
            $Usage.Currency = $Currency
        }
        $UsageData += $Usage
    }

    if ([int]($UsageData | measure -Sum -Property cost).Sum -gt $skipSize)
    {
        $RGCost.SubscriptionID = $Record.SubscriptionRG.SubscriptionId
        $RGCost.Subscription = $Record.SubscriptionRG.Subscription
        $RGCost.ResourceGroupName = $Record.SubscriptionRG.ResourceGroupName
        $RGCost.Link = $Record.SubscriptionRG.Link
        $RGCost.Location = $Record.SubscriptionRG.Location
        $RGCost.Owner= $Record.SubscriptionRG.Owner
        $RGCost.OwnerName= $Record.SubscriptionRG.Owner.Split('@')[0]
        
        $RGCost.Solution = $Record.SubscriptionRG.Solution
        $RGCost.OwnerSolution= ($RGCost.OwnerName + "-" + $RGCost.Solution)
        $RGCost.UsageData = $UsageData
        $RGCost.TotalCost = [int]($UsageData | measure -Sum -Property cost).Sum
        $RGCosts += $RGCost
    }
    else
    {
        write-verbose ("Record skipped cost less than $skipSize" )
    }
}
write-verbose ("Records created "  + [string]$RGCosts.Count)


$PeopleCosts = @()
$PeopleGrouping = ($RGCosts | group Owner)

foreach ($costGroup in $PeopleGrouping )
{
     $PeopleCost = '' | select Owner,OwnerName, UsageData ,Currency, Total
     $PeopleCost.Owner =  $costGroup.Name
     $PeopleCost.OwnerName = $costGroup.name.Split('@')[0]
     $UsageGrouping = $costGroup.Group.usagedata.date | select -Unique
     $PeopleCost.currency =  $costGroup.Group.usagedata.currency[0]
     $DailyTotals = @()
     foreach ($UsageDate in $UsageGrouping)
     {
        $DailyTotal = '' | select 'CostTotal','Currency','Date'
        $DailyRecords = @(@($costGroup.Group.usagedata) |  ? {$_.date -eq $UsageDate})
        $DailyTotal.Date = $UsageDate
        $DailyTotal.Currency = $DailyRecords[0].currency
        $DailyTotal.costTotal =  ($DailyRecords | measure cost -Sum).sum
        $DailyTotals += $DailyTotal
     }
     
     $PeopleCost.UsageData = $DailyTotals
     $PeopleCost.Total = ($DailyTotals | measure costtotal -Sum).sum
        
     $PeopleCosts+= $PeopleCost 
}



$rpt = @()
$tabarray = @('Stacked Charts','Costs','Resource Groups','Missing Tags')
if ($YouLogoHereURLString -ne $null)
{
    $rpt += Get-HTMLOpenPage -TitleText "Azure Cost Breakdown" -LeftLogoName Corporate -RightLogoName Alternate 
}
else
{
    $rpt += Get-HTMLOpenPage -TitleText "Azure Cost Breakdown" -LeftLogoName Corporate -RightLogoName Alternate 
}

$rpt += Get-HTMLTabHeader -TabNames $tabarray 
$rpt += get-htmltabcontentopen -TabName 'Stacked Charts' -tabheading ('Stacked Spend Charts')


        $StackChartObject1 = Get-HTMLStackedChartObject
        $StackChartObject1.Size.Width =1400
        $StackChartObject1.Size.Height = 600
        $StackChartObject1.DataDefinition.DataSetArrayData = 'UsageData'
        $StackChartObject1.DataDefinition.DataSetArrayDataColumn = 'Cost'
        $StackChartObject1.DataDefinition.DataSetArrayName ='OwnerSolution'
        $StackChartObject1.DataDefinition.DataSetArrayXLabels = 'Date'
        $StackChartObject1.ChartStyle.hover.mode = 'dataset'
        $StackChartObject1.ChartStyle.hover.intersect = 'true'
        $StackChartObject1.ChartStyle.tooltips.mode = 'dataset'
        $StackChartObject1.ChartStyle.tooltips.intersect = 'true'

        $StackChartObject2 = Get-HTMLStackedChartObject
        $StackChartObject2.Size.Width =1400
        $StackChartObject2.Size.Height = 600
        $StackChartObject2.DataDefinition.DataSetArrayData = 'UsageData'
        $StackChartObject2.DataDefinition.DataSetArrayDataColumn = 'CostTotal'
        $StackChartObject2.DataDefinition.DataSetArrayName ='OwnerName'
        $StackChartObject2.DataDefinition.DataSetArrayXLabels = 'Date'
        $StackChartObject2.ChartStyle.hover.mode = 'x'
        $StackChartObject2.ChartStyle.hover.intersect = 'false'
        $StackChartObject2.ChartStyle.tooltips.mode = 'x'
        $StackChartObject2.ChartStyle.tooltips.intersect = 'false'
        




        $rpt += Get-HTMLContentOpen -HeaderText "By Person"
            $rpt += Get-HTMLStackedChart -ChartObject $StackChartObject2 -DataSet ($PeopleCosts | sort total -Descending)
        $rpt += Get-HTMLContentclose


        $rpt += Get-HTMLContentOpen -HeaderText "Person & Solution Cost"
            $rpt += Get-HTMLStackedChart -ChartObject $StackChartObject1 -DataSet $RGCosts
        $rpt += Get-HTMLContentclose



$rpt += get-htmltabcontentclose



$rpt += get-htmltabcontentopen -TabName 'Resource Groups' -tabheading ('Resource Groups')

    if (!$AutoKeyName) 
    {
        $Pie1 = $SubscriptionRGs| group $KeyNames[0]
        $Pie2 = $SubscriptionRGs| group $KeyNames[1]

        $Pie1Object = Get-HTMLPieChartObject -ColorScheme Random
        $Pie2Object = Get-HTMLPieChartObject -ColorScheme Generated


        $rpt += Get-HTMLContentOpen -HeaderText "Pie Charts"
               $rpt += Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 2
                   $rpt += Get-HTMLPieChart -ChartObject $Pie1Object  -DataSet $Pie1
               $rpt += Get-HTMLColumnClose
               $rpt += Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 2
                   $rpt += Get-HTMLPieChart -ChartObject $Pie2Object -DataSet $Pie2
               $rpt += Get-HTMLColumnClose
        $rpt += Get-HTMLContentclose
    }
    	$rpt += Get-HTMLContentOpen -HeaderText "Complete List"
		$rpt += Get-HTMLContentdatatable -ArrayOfObjects ( $SubscriptionRGs  | select ResourceGroupName, Location, Link,Owner,Solution)
	$rpt += Get-HTMLContentClose 


$rpt += get-htmltabcontentclose


$rpt += get-htmltabcontentopen -TabName 'Costs' -tabheading ('Costs')

        $Pie3 =  $RGCosts | group Owner  | select Name, @{ Name = 'Count'; Expression = { ($_.Group | Measure-Object -Property TotalCost -Sum).Sum } }
        $Pie4 =  $RGCosts | group Solution  | select Name, @{ Name = 'Count'; Expression = { ($_.Group | Measure-Object -Property TotalCost -Sum).Sum } }

        $Pie3Object = Get-HTMLPieChartObject -ColorScheme Random
        $Pie4Object = Get-HTMLPieChartObject -ColorScheme Generated

        $rpt += Get-HTMLContentOpen -HeaderText "Pie Charts"
               $rpt += Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 2
                   $rpt += Get-HTMLPieChart -ChartObject $Pie3Object  -DataSet $Pie3
               $rpt += Get-HTMLColumnClose
               $rpt += Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 2
                   $rpt += Get-HTMLPieChart -ChartObject $Pie4Object -DataSet $Pie4
               $rpt += Get-HTMLColumnClose
        $rpt += Get-HTMLContentclose
        
        
	$rpt += Get-HTMLContentOpen -HeaderText "Complete List"
		$rpt += Get-HTMLContentdatatable -ArrayOfObjects (  $RGCosts   | select ResourceGroupName, Location, Link,Owner,Solution, TotalCost)
	$rpt += Get-HTMLContentClose 



$rpt += get-htmltabcontentclose


#Missing tags

$rpt += get-htmltabcontentopen -TabName 'Missing Tags' -tabheading ('Missing Tags')

  

    $rpt += Get-HTMLContentOpen -HeaderText "Resource Groups Missing Tags"
		$rpt += Get-HTMLContenttable -ArrayOfObjects ( $SubscriptionRGs  | ? {$_.owner -eq $null -or $_.solution -eq $null} |  select ResourceGroupName, Location, Link,Owner,Solution)
	$rpt += Get-HTMLContentClose 


$rpt += get-htmltabcontentclose

$rpt += Get-HTMLClosePage  

if ($ReportOutputPath -ne $null)
{
    $ReportFile  =  Save-HTMLReport -ReportContent $rpt -ReportName CostBreakdown -ReportPath $ReportOutputPath 
}
else
{
    Save-HTMLReport -ShowReport -ReportContent $rpt -ReportName CostBreakdown
}


$SendTo =  $PeopleCosts # | ? {$_.Owner -like $EmailDomain }
foreach ($People in $SendTo  )
{
    $DailyAvg = [math]::round(($People.UsageData | measure CostTotal -Average).Average,0)
    $DailyMax = [math]::round(($People.UsageData | measure CostTotal -Maximum).Maximum,0)
    $CostSum = [math]::round(($People.UsageData | measure CostTotal -Sum).sum,0)
    

    write-host "sending $($People.Owner) $($CostSum ) $($DailyMax)  $($DailyAvg)"
    $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)
    $Mail.To = $People.Owner
    $Mail.Subject = "Total: $('$')$($CostSum) | Daily Max: $('$')$($DailyMax)  Avg: $('$')$($DailyAvg)"
    $Mail.Body ="Please find attached Azure Spend Report.  "
    #write-host  "Total: $('$')$($CostSum) | Daily Max: $('$')$($DailyMax)  Avg: $('$')$($DailyAvg)"
    $Mail.Attachments.add($ReportFile)

    #$Mail.Send()
    
}

