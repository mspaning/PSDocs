#
# Unit tests for core PSDocs functionality
#

[CmdletBinding()]
param (

)

# Setup error handling
$ErrorActionPreference = 'Stop';
Set-StrictMode -Version latest;

# Setup tests paths
$rootPath = $PWD;

Import-Module (Join-Path -Path $rootPath -ChildPath out/modules/PSDocs) -Force;

$here = (Resolve-Path $PSScriptRoot).Path;
$outputPath = Join-Path -Path $rootPath -ChildPath out/tests/PSDocs.Tests/OpenXml;
Remove-Item -Path $outputPath -Force -Recurse -Confirm:$False -ErrorAction Ignore;
$Null = New-Item -Path $outputPath -ItemType Directory -Force;

Describe 'Invoke-PSDocument' -Tag 'OpenXml' {

    Context 'With -OutputFormat OpenXml' {

        It 'Should match name' {
            # Only generate documents for the named document
            Invoke-PSDocument -Path tests/PSDocs.Tests/ -OutputPath $outputPath -Name FromFileTest1 -OutputFormat OpenXml;
            $outputDoc = "$outputPath\FromFileTest1.docx";
            Test-Path -Path $outputDoc | Should -Be $True;
        }
    }
}
