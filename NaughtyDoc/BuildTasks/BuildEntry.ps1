#!/usr/bin/env pwsh
param($ProjectName, $ProjectDir, $TargetDir)
Pwsh -File "$($ProjectDir)\BuildTasks\HideBuild.ps1" $TargetDir test.docm "$true"