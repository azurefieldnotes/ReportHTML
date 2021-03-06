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

####### Example 1 ########
$rpt = @()
$rpt += Get-HtmlOpen -TitleText ($ReportName + "Example 1")
$rpt += Get-HtmlContentOpen -HeaderText "Virtual Machines"
$rpt += Get-HtmlContentTable $RMVMArray
$rpt += Get-HtmlContentClose 
$rpt += Get-HtmlClose

Test-Report -TestName Example1

####### Example 2 ########
$rpt = @()
$rpt += Get-HtmlOpen -TitleText  ($ReportName + "Example 2")
$rpt += Get-HtmlContentOpen -HeaderText "Virtual Machines"
$rpt += Get-HtmlContentTable ($RMVMArray | select Location, ResourceGroup, Name, Size,PowerState,  DataDiskCount, ImageSKU ) -GroupBy Location
$rpt += Get-HtmlContentClose 
$rpt += Get-HtmlClose

Test-Report -TestName Example2

####### Example 3 ########
$rpt = @()
$rpt += Get-HtmlOpen -TitleText ($ReportName + "Example 3")
$rpt += Get-HtmlContentOpen -HeaderText "Summary Information" 
$rpt += Get-HtmlContenttext -Heading "Total VMs" -Detail ( $RMVMArray.Count)
$rpt += Get-HtmlContenttext -Heading "VM Power State" -Detail ("Running " + ($RMVMArray | ? {$_.PowerState -eq 'Running'} | measure ).count + " / Deallocated " + ($RMVMArray | ? {$_.PowerState -eq 'Deallocated'} | measure ).count)
$rpt += Get-HtmlContenttext -Heading "Total Data Disks" -Detail $RMVMArray.datadisks.count
$rpt += Get-HtmlContentClose 
$rpt += Get-HtmlContentOpen -HeaderText "VM Size Summary" -IsHidden
$rpt += Get-HtmlContenttable ($RMVMArray | group size | select Name, Count | sort count -Descending ) -Fixed
$rpt += Get-HtmlContentClose 
$rpt += Get-HtmlContentOpen -HeaderText "Virtual Machines" -IsHidden
$rpt += Get-HtmlContentTable ($RMVMArray | select Location, ResourceGroup, Name, Size,PowerState,  DataDiskCount, ImageSKU ) -GroupBy Location
$rpt += Get-HtmlContentClose 
$rpt += Get-HtmlClose

Test-Report -TestName Example3

####### Example 4 ########
$rpt = @()
$rpt += Get-HtmlOpen -TitleText ($ReportName + "Example 4")
$rpt += Get-HtmlContentOpen -HeaderText "Summary Information" 
$rpt += Get-HtmlContenttext -Heading "Total VMs" -Detail ( $RMVMArray.Count)
$rpt += Get-HtmlContenttext -Heading "VM Power State" -Detail ("Running " + ($RMVMArray | ? {$_.PowerState -eq 'Running'} | measure ).count + " / Deallocated " + ($RMVMArray | ? {$_.PowerState -eq 'Deallocated'} | measure ).count)
$rpt += Get-HtmlContenttext -Heading "Total Data Disks" -Detail $RMVMArray.datadisks.count
$rpt += Get-HtmlContentClose 
$rpt += Get-HtmlContentOpen -HeaderText "VM Size Summary" -IsHidden
$rpt += Get-HtmlContenttable ($RMVMArray | group size | select Name, Count | sort count -Descending ) -Fixed
$rpt += Get-HtmlContentClose 
$rpt += Get-HtmlContentOpen -HeaderText "Virtual Machines by location" -IsHidden -BackgroundShade 2
foreach ($Group in ($RMVMArray | select Location, ResourceGroup, Name, Size,PowerState,  DataDiskCount, ImageSKU | group location ) ) {
	$rpt += Get-HtmlContentOpen -HeaderText ("Virtual Machines for location '" + $group.Name +"'") -IsHidden -BackgroundShade 1
	$rpt += Get-HtmlContentTable ($Group.Group | select PowerState,ResourceGroup, Name, Size,  DataDiskCount, ImageSKU ) -Fixed 
	$rpt += Get-HtmlContentClose 
}
$rpt += Get-HtmlContentClose 
$rpt += Get-HtmlClose

Test-Report -TestName Example4


####### Example 5 ########
$rpt = @()
$rpt += Get-HtmlOpen -TitleText ($ReportName + "Example 5")
$rpt += Get-HtmlContentOpen -HeaderText "Summary Information" -BackgroundShade 1
$rpt += Get-HtmlContenttext -Heading "Total VMs" -Detail ( $RMVMArray.Count)
$rpt += Get-HtmlContenttext -Heading "VM Power State" -Detail ("Running " + ($RMVMArray | ? {$_.PowerState -eq 'Running'} | measure ).count + " / Deallocated " + ($RMVMArray | ? {$_.PowerState -eq 'Deallocated'} | measure ).count)
$rpt += Get-HtmlContenttext -Heading "Total Data Disks" -Detail $RMVMArray.datadisks.count
$rpt += Get-HtmlContentClose 
$rpt += Get-HtmlContentOpen -HeaderText "VM Size Summary" -IsHidden -BackgroundShade 1
$rpt += Get-HtmlContenttable ($RMVMArray | group size | select Name, Count | sort count -Descending ) -Fixed
$rpt += Get-HtmlContentClose 
$rpt += Get-HtmlContentOpen -HeaderText "Virtual Machines by location" -BackgroundShade 3
foreach ($Group in ($RMVMArray | select Location, ResourceGroup, Name, Size,PowerState,  DataDiskCount, ImageSKU | group location ) ) {
	$PowerState = $Group.Group | group PowerState 
	$rpt += Get-HtmlContentOpen -HeaderText ("Virtual Machines for location '" + $group.Name +"' - "  + $Group.Group.Count + " VMs") -IsHidden -BackgroundShade 2
		if (($PowerState | ? {$_.name -eq 'Running'})) {
			$rpt += Get-HtmlContentOpen -HeaderText ("Running Virtual Machines") -BackgroundShade 1
			$rpt += Get-HtmlContentTable ($Group.Group | where {$_.PowerState -eq "Running"} | select ResourceGroup, Name, Size,  DataDiskCount, ImageSKU  ) -Fixed 
			$rpt += Get-HtmlContentClose 
		}
		if (($PowerState | ? {$_.name -eq 'Deallocated'})) {
			$rpt += Get-HtmlContentOpen -HeaderText ("Deallocated") -BackgroundShade 1 -IsHidden
			$rpt += Get-HtmlContentTable ($Group.Group  | where {$_.PowerState -eq "Deallocated"} | select ResourceGroup, Name, Size,  DataDiskCount, ImageSKU)-Fixed 
			$rpt += Get-HtmlContentClose 
		}
	$rpt += Get-HtmlContentClose 
}
$rpt += Get-HtmlContentClose 
$rpt += Get-HtmlClose

Test-Report -TestName Example5

Invoke-Item $ReportOutputPath
