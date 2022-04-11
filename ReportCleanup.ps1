<#
    .Synopsis
        Moves SSRS reports to an archive project based on when the report was last ran and how many times it was ran.
        
    .Description
        Uses the built-in SSRS database to get the report name, path, times ran, and last run date. It will then
        deploy the report to the SSRS archive, move the rdl file to the archive visual studio folder, add the data file
        to the archive project file, and remove the report from the previous project file.
    #>

#Parameters for the dateadd function in SQL using today's date. #Lookup SQL DATEADD for all parameter options.
$interval = 'y'
$timeFrame = -1

#The number of times the report has been ran.
$timesRan = 5

$reportServer = 'http://LBSSQL/ReportS'
$archiveFolder = '/Archive'
$projectsFolder = '\\lbscares.com\storage\group_documents\jmunro\Visual Studio 2010\Projects\'

#The archive project file we'll move reports to. It has an extra base folder.
$projectFilePath = '\\lbscares.com\storage\group_documents\jmunro\Visual Studio 2010\Projects\archive\Archive\Archive\archive.rptproj'

$projectFile = get-content -path $projectFilePath

#Using the report server database to pull a list of all reports, the number of times they were ran, and when they were last ran.
#Cat.Type = 2 is used to filter only reports rather than folders.
$queryParams = @{
    Hostname = 'LBSSQL'
    ServerInstance = 'LBSSQL\'
    Database = 'ReportServer'
    Query = "use ReportServer
        select  cat.Name, cat.path, count(TimeStart) as TimesRan, cast(max(ex.timestart)as date) as lastrun
        from ExecutionLog as ex
        right join Catalog as cat on ex.reportid = cat.itemid
        where path not like '%Archive%'
        and path not like '%All Staff%'
        and cat.type = 2
        group by cat.Name, cat.Path
        having count(TimeStart) < $timesRan and (
	        cast(max(ex.timestart)as date) < dateadd($interval,$timeFrame,getdate()) 
	        or (cast(max(ex.timestart)as date) is null)
        )
        order by lastrun"
}

$queryResults = Invoke-Sqlcmd @queryParams

#The report path from SQL uses the forward slash so it needs to be replaced.
#There's an extra folder in the windows path so we need an additional base folder
#We the need the path to the project file of the report currently in the loop so the report name gets replaced with the project file name.
$queryResults| foreach {
  $fixPathSlash = ($_.path).replace('/','\') 
  $fixPathFolder = ($_.path).split('/')[1]
  $finalPath = $projectsFolder+$fixPathFolder+$fixPathSlash
  $currentProjectPath = ($finalpath).replace($_.Name, "$fixPathFolder.rptproj")

   
  #Writes the report to the archive ssrs folder, removes it from the old folder, and moves the report definition file to the archive folder.
  Write-RsCatalogItem -ReportServerUri $reportServer -path "$($finalPath).rdl" -RsFolder $archiveFolder -Verbose
  Remove-RsCatalogItem -ReportServerUri $reportServer -rsitem $_.path -Verbose -Confirm $false
  move-item -Path "$($finalPath).rdl" -Destination "$($projectsFolder)archive\archive\archive\" -Verbose

  #This is the text that is added to the archive report project file and is removed from thom the current project file. 
  #In order to get it in the proper spot we need to tell it what line number to enter the text.
  $addText = "
  <ProjectItem>
    <Name>$($_.Name).rdl</Name>
    <FullPath>$($_.Name).rdl</FullPath>
  </ProjectItem>"
  $projectFile[$linenumber+18] += $addText
  
  #The files are often still in use so we need a basic while loop to check if the file is available.
  #Keep the current project file the same except where it matches the addtext variable.
  $completed = $false
  while (-not $completed) {
    If (get-content -path $currentProjectPath) {
      get-content -path $currentProjectPath | where-object {$_ -notmatch $addtext} | set-content -path $currentProjectPath
      $completed = $True
    }
    else {
      write-verbose "File is still in use."
      start-sleep '30 seconds'
    }
  }
}

#The files are often still in use so I used a basic while loop.
#Set the updated arhive project file after going through the for loop.
$completed = $false
while (-not $completed) {
  If (get-content -path $projectFilePath) {
    $projectFile | set-content -path $projectFilePath
    $completed = $True
  }
  else {
    write-verbose "File is still in use."
    start-sleep '30 seconds'
  }
}




