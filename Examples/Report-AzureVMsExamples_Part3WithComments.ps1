param (
	$ReportOutputPath
)

Import-Module ReportHtml
Get-Command -Module ReportHtml

if (!$ReportOutputPath) 
{
	$ReportOutputPath = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
} 
$ReportName = "Azure VMs"

# see if we already have a session. If we don't don't re-authN
if (!$AzureRMAccount.Context.Tenant) {
    $AzureRMAccount = Add-AzureRmAccount 		
}

# Get arrary of VMs from ARM
$RMVMs = get-azurermvm 

$RMVMArray = @() ; $TotalVMs = $RMVMs.Count; $i =1 
# Loop through VMs
foreach ($vm in $RMVMs)
{
  # Tracking progress
  Write-Progress -PercentComplete ($i / $TotalVMs * 100) -Activity "Building VM array" -CurrentOperation  ($vm.Name + " in resource group " + $vm.ResourceGroupName)
    
  # Get VM Status (for Power State)
  $vmStatus = Get-AzurermVM -Name $vm.Name -ResourceGroupName $vm.ResourceGroupName -Status

  # Generate Array
  $RMVMArray += New-Object PSObject -Property @{`

    # Collect Properties
   	ResourceGroup = $vm.ResourceGroupName
	ID = $VM.id
	Name = $vm.Name;
    PowerState = (get-culture).TextInfo.ToTitleCase(($vmStatus.statuses)[1].code.split("/")[1]);
    Location = $vm.Location;
    Tags = $vm.Tags
    Size = $vm.HardwareProfile.VmSize;
    ImageSKU = $vm.StorageProfile.ImageReference.Sku;
    OSType = $vm.StorageProfile.OsDisk.OsType;
    OSDiskSizeGB = $vm.StorageProfile.OsDisk.DiskSizeGB;
    DataDiskCount = $vm.StorageProfile.DataDisks.Count;
    DataDisks = $vm.StorageProfile.DataDisks;
    }
	$i++
}
  
Function Test-Report 
{
	param (
		$TestName
	)
	$rptFile = join-path $ReportOutputPath ($ReportName.replace(" ","") + "-$TestName" + ".mht")
	$rpt | Set-Content -Path $rptFile -Force
	Invoke-Item $rptFile
	sleep 1
}


####### Example 9 ########

# First we create a PieChart Object and load it into a variable 
$PieChartObject = Create-HTMLPieChartObject

# have a look at what is in this object.  
$PieChartObject 

# let's set one property
$PieChartObject.Title = "VMs Power State"

$rpt = @()
$rpt += Get-HtmlOpen -TitleText  ($ReportName + " Example 9")
$rpt += Get-HtmlContentOpen -HeaderText "Chart Series"
$rpt += Create-HTMLPieChart -PieChartObject $PieChartObject -PieChartData ($RMVMArray | group powerstate)
$rpt += Get-HtmlContentClose 
$rpt += Get-HtmlClose

Test-Report -TestName Example9

####### Example 10 ########	 	 
$PieChartObject = Create-HTMLPieChartObject	 	 
$PieChartObject.Title = "VMs Sizes Deployed"

# There is a lot of data so let's make the pie chart a little bigger and explode the largest value	 	 
$PieChartObject.Size.Height = 600	 	 
$PieChartObject.Size.Width = 600
$PieChartObject.ChartStyle.ExplodeMaxValue = $true	 	 

$rpt = @()	 	 
$rpt += Get-HtmlOpen -TitleText ($ReportName + " Example 10")	 	 
$rpt += Get-HtmlContentOpen -HeaderText "Chart Series"	 	 

# To summarize the data I have simply changed the group by property to size 	 	 
$rpt += Create-HTMLPieChart -PieChartObject $PieChartObject -PieChartData ($RMVMArray | group size)	 	 
$rpt += Get-HtmlContentClose 	 	 
$rpt += Get-HtmlClose	 	 
Test-Report -TestName Example10	 


Invoke-Item $ReportOutputPath
