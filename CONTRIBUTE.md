## How to contribute
Thanks for your interests and effort to contribute to this module!!! These are basic steps for how to contribute:smile:

#### Start from addng an issue
Regardleess you fix the issue or add new function, please add an issue first so that everyone can aware what's the current issues/challenges.

#### Fork the repo
Fork the repo by click "Fork" button on right top corner, which brings entire repo to your GitHub account.

#### Add Branch as Issue number
Add a branch to work on. Please name the branch as issue number so that we can easily relate a pull request to an issue.

#### Pull Request
Once you fixed the issue or added new function, please make pull request.

For the new function, please use the following template.

- Use {Verb}-Crm{Noun} to name function
- Specify Mandatory and Position for parameters (and more if you need)
- Leave $conn parameter as it is to make all the functions experience same
- Please add enough help and examples
- Use Upper case for Public Parameter, and lowercase for local parameter

```powershell
function Verb-CrmNoun{
<#
 .SYNOPSIS
 Brief overview of the function

 .DESCRIPTION
 Function desctiprion here
 
 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER Param1
 Param1 detail
 
 .PARAMETER Param2
 Param2 detail
 
 .EXAMPLE 
 Here is example
#>
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$Param1,
        [parameter(Mandatory=$false, Position=2)]
        [int]$Param2
    )
      $conn = VerifyCrmConnectionParam $conn
      ## Do work
}
```
