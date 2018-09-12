#requires -Version 3.0 -Modules Pester
<#
        .SYNOPSIS
        This is a PowerShell Unit Test file for a PSModule's functions' help text.

        .DESCRIPTION
        You need a unit test framework such as Pester to run PowerShell Unit tests. 
        You can download Pester from http://go.microsoft.com/fwlink/?LinkID=534084

        .NOTES
        // Copyright (c) Microsoft Corporation. All rights reserved.
        // Licensed under the MIT license.

        # Who             What            When            Why
        # timdunn         v1.0.0          2018-09-07      Created this because I'm a masochist.

        Drop this fine in the PSModule's root folder. Invoke-Pester will test it
        against all the functions the root module exports.
#>

# disallow other parameters, support CommonParameters
[CmdletBinding()]param()

[string]$moduleName = Split-Path -Path $PSScriptRoot -Leaf
Import-Module -Name $PSScriptRoot -Force -Global

Describe "$moduleName function comment-based help" `
{
    <#
            .SYNOPSIS
            Make sure all functions have adequate comment-based-help text.
    #>

    Get-Command -Module $moduleName -CommandType Function |
    ForEach-Object -Process `
    {
        # Generate a Context{} scriptblock for each function

        [Management.Automation.FunctionInfo]$cmdInfo = $_
        [string]$functionName = $cmdInfo.name
        
        $helpData = Get-Help -Name $moduleName\$functionName

        # these are the parameters the function defines
        [string[]]$functionDefinedParameterNames = $_.parameters.Values |
        Where-Object -FilterScript `
        {
            # CommonParameters define their Position attribute to -2147483648
            $_.attributes.Position -ne [int]::MinValue
        } |
        Select-Object -ExpandProperty Name


        Context "$functionName comment-based help" `
        {
            #region comment-based-help sections besides .PARAMETER

            It "$functionName comment-based help has a .SYNOPSIS section" `
            {
                # if synopsis is not defined, it is set to the multiline string:
                # '
                # Function-Name
                # '

                $synopsis = $helpData.Synopsis -replace '^\s+' -replace '\s+$'

                if ( $synopsis -eq $functionName )
                {
                    # by trimming the leading and trailing whitespace, we can
                    # detect if the comment-based help has a .SYNOPSIS section.
                    # And if the .SYNOPSIS section _is_ the function's name,
                    # then we have another problem...

                    Write-Verbose -Message 'Synopsis: '
                }
                else
                {
                    Write-Verbose -Message "Synopsis: '$synopsis'"
                }

                $Synopsis |
                Should Not Be $functionName
            }


            It "$functionName comment-based help has a .DESCRIPTION section" `
            {
                # each function must have a .DESCRIPTION comment-based help section

                $description = $helpData.description

                Write-Verbose -Message "Description: $description"

                $description |
                Should Not Be $null
            }


            It "$functionName comment-based help has an .INPUTS section" `
            {
                # each function must have an .INPUTS comment-based help section

                $inputs = $helpData.inputTypes.inputType.type.name

                Write-Verbose -Message "Inputs: $inputs"

                $inputs |
                Should Not Be $null
            }

            # we're going to test for the .PARAMETERs being in .EXAMPLEs later
            # so let's save the example code.
            [string[]]$exampleCode = $helpData.examples.example.code

            It "$functionName comment-based help has an .EXAMPLES section" `
            {
                # each function must have an .EXAMPLE comment-based help section

                if ( $exampleCode )
                {
                    $exampleCode |
                    Write-Verbose
                }
                else
                {
                    Write-Verbose -Message 'Examples:'
                }

                $exampleCode |
                Should Not Be $null
            }

            #endregion comment-based-help sections besides .PARAMETER
            #region comment-based-help .PARAMETER sections

            # comment-based help will always have a .PARAMETER section for each
            # param() command line parameter.
            $helpData.parameters.parameter |
            Where-Object -Property Name -In -Value $functionDefinedParameterNames |
            ForEach-Object -Process `
            {
                $parameterName = $_.Name
            
                It "$functionName's -$parameterName parameter has help documentation" `
                {
                    # the default comment-based help parameter data will
                    # not have a .description property value

                    $description = $_.description

                    Write-Verbose -Message "Description: $description"

                    $description |
                    Should Not Be $null
                }

                It "$functionName's -$parameterName parameter is either required, or has a default value" `
                {
                    # this isn't so much user help as it is code hygiene.
                    # PowerShell defaults all variables to $false, $null, 0
                    # or the empty set, depending on parameter type. However
                    # good code hygiene is to always initialize variables if
                    # the user does not set them when calling the function.

                    $required     = $_.Required
                    $defaultValue = $_.DefaultValue

                    Write-Verbose -Message "Required: $required"
                    Write-Verbose -Message "DefaultValue: '$defaultValue'"

                    $required -or $defaultValue |
                    Should Be $true
                }

                It "$functionName's -$parameterName parameter is in an example" `
                {
                    # EVERY parameter should be in an .EXAMPLE section. EVERY. ONE.
                    # note that we are not including CommonParameters in this loop

                    $example = $exampleCode -match "\s-$parameterName\s" |
                    Select-Object -First 1

                    Write-Verbose -Message "Example: $example"

                    $example |
                    Should Not Be $null

                    $example |
                    Should Not Be $false
                }

            } # $helpData.Parameters.Parameter | ForEach-Object

            #region comment-based-help .PARAMETER sections 

        } # Context

    } # Get-Command | ForEach-Object

    #> # Describe Help
}


