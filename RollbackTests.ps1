$VerbosePreference = "Continue"

# $true to run without applying changes
$TestMode = $true
# TFS/VSTS team project collection url 
$TargetUrl = "http://host:port/tfs/collection"
# Test project name
$TargetProject = "Team Project"
# Work item IDs using the following format (1,2,3)
$TargetIds = "(1,2,3)"
# Target date for the rollback
$TargetDate = Get-Date -Year 1985 -Month 10 -Day 26 -Hour 09 -Minute 00 -Second 00
# TFS API DLLs location
$AssemblyFolder = "${Env:ProgramFiles(x86)}\Microsoft Visual Studio 14.0\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer"

[System.Reflection.Assembly]::LoadFrom("$AssemblyFolder\Microsoft.TeamFoundation.Client.dll") 
[System.Reflection.Assembly]::LoadFrom("$AssemblyFolder\Microsoft.TeamFoundation.TestManagement.Client.dll") 

$TargetTFS = New-Object Microsoft.TeamFoundation.Client.TfsTeamProjectCollection(New-Object System.Uri($TargetUrl))
#For use with VSTS:  $TargetTFS.ClientCredentials = New-Object Microsoft.TeamFoundation.Client.TfsClientCredentials(New-Object Microsoft.TeamFoundation.Client.BasicAuthCredential(New-Object System.Net.NetworkCredential("", "pat")))
$TargetTFS.Authenticate()

$TargetTestMgmt = $TargetTFS.GetService([Microsoft.TeamFoundation.TestManagement.Client.ITestManagementService])
$TargetTeamProject = $TargetTestMgmt.GetTeamProject($TargetProject)
$WorkItems = $TargetTeamProject.WitProject.Store.Query("SELECT * FROM WorkItems WHERE [System.Id] IN $TargetIds")

foreach($WorkItem in $WorkItems)
{
    $GoodRevision = $WorkItem.Revisions | ? { $_.Fields["System.ChangedDate"].Value -lt $TargetDate } | select -Last 1

    if ($GoodRevision)
    {
        if ($WorkItem.Rev -eq $GoodRevision.Fields["System.Rev"].Value)
        {
            Write-Verbose "Test Case $($WorkItem.Id) has no revision after $TargetDate"
        }
        else
        {
            Write-Verbose "Test Case $($WorkItem.Id) will be updated with revision $($GoodRevision.Fields["System.Rev"].Value) ($($GoodRevision.Fields["System.ChangedDate"].Value))"

            $TestCase = $TargetTeamProject.TestCases.Find($WorkItem.Id)

            $TestCase.Description = $GoodRevision.Fields["System.Description"].Value
            ($TestCase.CustomFields | ? { $_.ReferenceName -eq "Microsoft.VSTS.TCM.Steps" }).Value = $GoodRevision.Fields["Microsoft.VSTS.TCM.Steps"].Value
            ($TestCase.CustomFields | ? { $_.ReferenceName -eq "Microsoft.VSTS.TCM.Parameters" }).Value = $GoodRevision.Fields["Microsoft.VSTS.TCM.Parameters"].Value
            ($TestCase.CustomFields | ? { $_.ReferenceName -eq "Microsoft.VSTS.TCM.LocalDataSource" }).Value = $GoodRevision.Fields["Microsoft.VSTS.TCM.LocalDataSource"].Value

            if (!$TestMode) { $TestCase.Save() }
        }
    }
    else
    {
        Write-Verbose "Test Case $($WorkItem.Id) has no revision before $TargetDate"
    }
}

