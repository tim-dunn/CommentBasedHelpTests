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
Import-Module -Name $PSScriptRoot -Force -Global -ErrorAction Stop

$psdPath = $PSCommandPath -replace '\.ps1$', '.psd1'
if ( Test-Path -Path $psdPath )
{
    $destination = "$env:TEMP\$( [guid]::NewGuid().ToString ).ps1"
    Copy-Item -Path $psdPath -Destination $destination
    [hashtable]$script:testSettings = . $destination
}


Describe "$moduleName function comment-based help" `
{
    <#
            .SYNOPSIS
            Make sure all functions have adequate comment-based-help text.
    #>

    [Collections.ArrayList]$script:parameterHelpText = @()

    [Management.Automation.FunctionInfo[]]$functions =
    Get-Command -Module $moduleName -CommandType Function

    foreach ( $cmdInfo in $functions)
    {
        # Generate a Context{} scriptblock for each function
        [string]$functionName = $cmdInfo.name

        [PSCustomObject[]]$helpData = Get-Help -Name $moduleName\$functionName

        # we're going to test for the .PARAMETERs being in .EXAMPLEs later
        # so let's save the example code.
        [string[]]$script:exampleCode = $helpData.examples.example.code

        # these are the parameters the function defines
        [string[]]$functionDefinedParameterNames = $cmdInfo.parameters.Values |
        Where-Object -FilterScript `
        {
            # CommonParameters define their Position attribute to -2147483648
            $_.attributes.Position -ne [int]::MinValue
        } |
        Select-Object -ExpandProperty Name


        Context "$functionName comment-based help" `
        {

            It "$functionName comment-based help has a  .SYNOPSIS section" `
            {
                # if synopsis is not defined, it is set to the multiline string:
                # '
                # Function-Name
                # '

                [string]$synopsis = $helpData.Synopsis -replace '^\s+' -replace '\s+$'

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


            It "$functionName comment-based help has a  .DESCRIPTION section" `
            {
                # each function must have a .DESCRIPTION comment-based help section

                [string]$description = $helpData.description

                Write-Verbose -Message "Description: $description"

                $description |
                Should Not Be $null
            }


            It "$functionName comment-based help has an .INPUTS section" `
            {
                # each function must have an .INPUTS comment-based help section

                [string]$inputs = $helpData.inputTypes.inputType.type.name

                Write-Verbose -Message "Inputs: $inputs"

                $inputs |
                Should Not Be $null
            }


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

        } #

        Context "$functionName .PARAMETER comment-based help" `
        {
            # comment-based help will always have a .PARAMETER section for each
            # param() command line parameter.
            $helpData.parameters.parameter |
            Where-Object -Property Name -In -Value $functionDefinedParameterNames |
            ForEach-Object -Process `
            {
                [string]$parameterName = $_.Name
            
                It "$functionName's -$parameterName parameter has help documentation" `
                {
                    # the default comment-based help parameter data will
                    # not have a .description property value

                    [string]$description = $_.description.Text

                    # cache the parameter text for consistency checking
                    $script:parameterHelpText += [PSCustomObject]@{
                        FunctionName  = $functionName
                        ParameterName = $parameterName
                        Description   = $description
                    }

                    Write-Verbose -Message "Description: $description"

                    $description |
                    Should Not Be $null
                }

                It "$functionName's -$parameterName parameter is  either required, or has a default value" `
                {
                    # this isn't so much user help as it is code hygiene.
                    # PowerShell defaults all variables to $false, $null, 0
                    # or the empty set, depending on parameter type. However
                    # good code hygiene is to always initialize variables if
                    # the user does not set them when calling the function.

                    [string]$required     = $_.Required
                    [string]$defaultValue = $_.DefaultValue

                    Write-Verbose -Message "Required: $required"
                    Write-Verbose -Message "DefaultValue: '$defaultValue'"

                    $required -or $defaultValue |
                    Should Be $true
                }

                It "$functionName's -$parameterName parameter is  in an example" `
                {
                    # EVERY parameter should be in an .EXAMPLE section. EVERY. ONE.
                    # note that we are not including CommonParameters in this loop

                    [string]$example = $script:exampleCode -match "\s-$parameterName\s" |
                    Select-Object -First 1

                    Write-Verbose -Message "Example: $example"

                    $example |
                    Should Not Be $null

                    $example |
                    Should Not Be $false
                }

            } # $helpData.Parameters.Parameter | ForEach-Object

        } 

    } # Get-Command | ForEach-Object


    Context "$moduleName .PARAMETER comment-based help consistency" `
    {
        # uncomment the next line to skip this set of tests. you won't hurt my feelings.
        # return

        filter Update-Description
        {
            # .SYNOPSIS
            # Dynamically rewrite .PARAMETER description text to ensure consistency.
            #
            # .PARAMETER ParameterName
            # Name of parameter for which to rewrite the description text. This Pester script
            # tests to ensure all text for a given parameter name is consistent across all
            # functions in the module, so some exceptions need to be made, such as when 
            # one file updates a file and another deletes it.
            #
            # .EXAMPLE
            # [string[]]$description = $parameterHelpText |
            # Where-Object -Property ParameterName -EQ -Value $ParameterName |
            # Update-Description -ParameterName $ParameterName |
            # Select-Object -ExpandProperty Description
            #
            # This will search apply the rewrite rules for the $ParameterName value in the
            # CommentBasedHelp.tests.psd1 ParameterDescriptionRewriteRules section for
            # the $parameterName value.
            #
            # .EXAMPLE
            # $description.Count | Should Be 1
            #
            # This tests for help documentaiton consistency for $ParameterName for all
            # functions in the PSModule

            param(
                [Parameter(
                        Mandatory,
                        HelpMessage = 'Parameter name for which to rewrite help text'
                )]
                [string]$ParameterName,


                [Parameter(
                        ValueFromPipeline,
                        ValueFromPipelineByPropertyName
                )]
                [Object[]]$InputObject
            )

            $InputObject |
            ForEach-Object -Process `
            {
                $_.Description =  $_.Description -replace '[\n\r]+', ' '

                if (
                    $Script:testSettings.ParameterDescriptionRewriteRules -and
                    (
                        $script:testSettings.ParameterDescriptionRewriteRules -is `
                        [hashtable] -and
                        $script:testSettings.ParameterDescriptionRewriteRules.ContainsKey( $ParameterName )
                    ) -or
                    (
                        $script:testSettings.ParameterDescriptionRewriteRules -is `
                        [Collections.Specialized.OrderedDictionary] -and
                        $script:testSettings.ParameterDescriptionRewriteRules.Contains( $ParameterName )
                    )
                )
                {
                    foreach ( $regex in $testSettings.ParameterDescriptionRewriteRules.$ParameterName )
                    {
                        if ( $regex.Search )
                        {
                            $_.Description  = $_.Description -replace $regex.Search, $regex.Replace
                        }
                    }
                }

                $_.Description =  $_.Description -replace '^\s+' -replace '\s+$'

                $_
            }

        } # filter Update-Description

        try
        {
            # use the Get-Something -Name target* | Where-Object Name -eq target to cut down on
            # needless exceptions

            [Collections.Specialized.OrderedDictionary]$parameterText =
            Get-Variable -ValueOnly -Scope Global -Name ModuleParameterText* -ErrorAction Stop |
            Where-Object -Property Name -EQ -Value ModuleParameterText

            if ( $parameterText.GetType().ToString() -ne 'System.Collections.Specialized.OrderedDictionary' )
            {
                # if it's the wrong type, re-initialize it
                throw "`$global:ModuleParameterText is not of type [System.Collections.Specialized.OrderedDictionary]"
            }
        }

        catch
        {
            # if it's not defined, declare it
            [Collections.Specialized.OrderedDictionary]$parameterText = @{}
        }

        finally
        {
            # add this module to the [ordered] [hashtable], now that we are sure it exists.
            $parameterText.$moduleName = [ordered] @{}        
        }

        [string[]]$parameterNames = $script:parameterHelpText.ParameterName |
        Sort-Object -Unique 

        foreach ( $_parameterName in $parameterNames )
        {
            It "$moduleName's functions with -$_parameterName have consistent help text." `
            {
                [PSCustomObject[]]$description = $script:parameterHelpText |
                Where-Object -FilterScript `
                {
                    # we're only looking at this parameter name
                    $_.ParameterName -eq $_parameterName -and

                    # ignore these functions for this parameter
                    $_.FunctionName -notin `
                    $script:testSettings.ExcludeParameterDescriptionConsistencyCheck.$_parameterName
                } |
                Update-Description -ParameterName $_parameterName |
                Group-Object Description

                $parameterText.$moduleName.$_parameterName = $description

                $description.Group |
                ForEach-Object -Process `
                {
                    Write-Verbose -Message "$($_.FunctionName) -$_parameterName is '$($_.Description)'."
                }

                $description.Count |
                Should Be 1
            }

        }

        # let's export the parameter help text data

        [Collections.Specialized.OrderedDictionary]$thisModuleParameterText = @{}

        foreach ( $_parameterName in $parameterText.$moduleName.Keys )
        {
            # we can't assign this directly because we can't modify the [hashtable] as we iterate over it
            $thisModuleParameterText.$_parameterName = $parameterText.$moduleName.$_parameterName.group
        }

        # so we create a temporary [hashtable], store the pertinent values, then save it as the data for this module.
        $parameterText.$moduleName = $thisModuleParameterText

        # and we export it to the global scope
        Set-Variable -Name ModuleParameterText -Value $parameterText -Scope Global -Force -Option AllScope

        # and tell the user what we're doing
        Write-Host -ForegroundColor Cyan -Object `
        "`$global:ModuleParameterText.'$moduleName' contains all -Parameter help text for this module."

    } # Context "$functionName .PARAMETER comment-based help consistency" `


    #> # Describe "$moduleName function comment-based help" 
}

# SIG # Begin signature block
# MIIcXgYJKoZIhvcNAQcCoIIcTzCCHEsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU499BxgT7HnGhxp7VGW0D/LWk
# PRaggheNMIIFFjCCA/6gAwIBAgIQBvCMmhZtT5VpGdBVvL78DzANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTE4MDgyOTAwMDAwMFoXDTIwMDEw
# MzEyMDAwMFowUzELMAkGA1UEBhMCVVMxCzAJBgNVBAgTAldBMREwDwYDVQQHEwhC
# ZWxsZXZ1ZTERMA8GA1UEChMIVGltIER1bm4xETAPBgNVBAMTCFRpbSBEdW5uMIIB
# IjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA1agFn56KigoW66r4WrzLD/ih
# f6SozX/XGXcQDXv4Ru5lba+L98p6m5AsfnSKTR4iJvT5pA4gZH+UyVuuxxKHXcnk
# pNUnEFhpumgkyCmP2dMDbBKxuyMT6jR/WsHar5IugW5+G/nBmwB9QCB805f4SQmB
# ob1gq8w+WAsNbY8yGIKSP4zKV5pB/5skTEv6UNkR58eZPOAI+3xqBo609RDSIHCt
# bYzSOPyKdo6iUy0NWN1vjEZ/X0RlTopM4FbBmVxdH1PeLqnGa5cw2BltgGPC+AEl
# wHxgLkRokcnEPiV7e79aqQtLne5yB6rnpGG2Q+/W22y4axep3AkuJ/vMVI5hyQID
# AQABo4IBxTCCAcEwHwYDVR0jBBgwFoAUWsS5eyoKo6XqcQPAYPkt9mV1DlgwHQYD
# VR0OBBYEFNHVUCqbfCNHU/OhZt+SjDkK5TpdMA4GA1UdDwEB/wQEAwIHgDATBgNV
# HSUEDDAKBggrBgEFBQcDAzB3BgNVHR8EcDBuMDWgM6Axhi9odHRwOi8vY3JsMy5k
# aWdpY2VydC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNybDA1oDOgMYYvaHR0cDov
# L2NybDQuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmwwTAYDVR0g
# BEUwQzA3BglghkgBhv1sAwEwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGln
# aWNlcnQuY29tL0NQUzAIBgZngQwBBAEwgYQGCCsGAQUFBwEBBHgwdjAkBggrBgEF
# BQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tME4GCCsGAQUFBzAChkJodHRw
# Oi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRTSEEyQXNzdXJlZElEQ29k
# ZVNpZ25pbmdDQS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAQEA
# NsxgWz7VOX7L8oJr7U1utb+FdoIfVerJG11LezE7YWj2vJZIoT038DghGQoSacbm
# An3WzRL60tXAeb6H6GDQqnYMgdUyNPJroRCv7F5TgR1xZIk18xC5kFgoDLTLWa5n
# IuN/AjO4uzp3QR9xxEDljYOaNEa8mZ6JrKzrYtxuIsk6ifTjAmPV18Q+JSM8U+S5
# emUoF/4PK8FAfO0XUvYQzU9PkBoFv+w8WV2en3SuWQKemzlq7ma9k/kflLS9mChS
# RdRoWNx+ngCcw3BTgiOqz1KV50YkqPR1dyNnSv5B+E7CVu9CN2x+6MO/BgVlkOBo
# abV0eqp5RzJPKgimDztMHDCCBTAwggQYoAMCAQICEAQJGBtf1btmdVNDtW+VUAgw
# DQYJKoZIhvcNAQELBQAwZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0
# IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNl
# cnQgQXNzdXJlZCBJRCBSb290IENBMB4XDTEzMTAyMjEyMDAwMFoXDTI4MTAyMjEy
# MDAwMFowcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcG
# A1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBB
# c3N1cmVkIElEIENvZGUgU2lnbmluZyBDQTCCASIwDQYJKoZIhvcNAQEBBQADggEP
# ADCCAQoCggEBAPjTsxx/DhGvZ3cH0wsxSRnP0PtFmbE620T1f+Wondsy13Hqdp0F
# LreP+pJDwKX5idQ3Gde2qvCchqXYJawOeSg6funRZ9PG+yknx9N7I5TkkSOWkHeC
# +aGEI2YSVDNQdLEoJrskacLCUvIUZ4qJRdQtoaPpiCwgla4cSocI3wz14k1gGL6q
# xLKucDFmM3E+rHCiq85/6XzLkqHlOzEcz+ryCuRXu0q16XTmK/5sy350OTYNkO/k
# tU6kqepqCquE86xnTrXE94zRICUj6whkPlKWwfIPEvTFjg/BougsUfdzvL2FsWKD
# c0GCB+Q4i2pzINAPZHM8np+mM6n9Gd8lk9ECAwEAAaOCAc0wggHJMBIGA1UdEwEB
# /wQIMAYBAf8CAQAwDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMD
# MHkGCCsGAQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNl
# cnQuY29tMEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20v
# RGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3J0MIGBBgNVHR8EejB4MDqgOKA2hjRo
# dHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0Eu
# Y3JsMDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1
# cmVkSURSb290Q0EuY3JsME8GA1UdIARIMEYwOAYKYIZIAYb9bAACBDAqMCgGCCsG
# AQUFBwIBFhxodHRwczovL3d3dy5kaWdpY2VydC5jb20vQ1BTMAoGCGCGSAGG/WwD
# MB0GA1UdDgQWBBRaxLl7KgqjpepxA8Bg+S32ZXUOWDAfBgNVHSMEGDAWgBRF66Kv
# 9JLLgjEtUYunpyGd823IDzANBgkqhkiG9w0BAQsFAAOCAQEAPuwNWiSz8yLRFcgs
# fCUpdqgdXRwtOhrE7zBh134LYP3DPQ/Er4v97yrfIFU3sOH20ZJ1D1G0bqWOWuJe
# JIFOEKTuP3GOYw4TS63XX0R58zYUBor3nEZOXP+QsRsHDpEV+7qvtVHCjSSuJMbH
# JyqhKSgaOnEoAjwukaPAJRHinBRHoXpoaK+bp1wgXNlxsQyPu6j4xRJon89Ay0BE
# pRPw5mQMJQhCMrI2iiQC/i9yfhzXSUWW6Fkd6fp0ZGuy62ZD2rOwjNXpDd32ASDO
# mTFjPQgaGLOBm0/GkxAG/AeB+ova+YJJ92JuoVP6EpQYhS6SkepobEQysmah5xik
# mmRR7zCCBmowggVSoAMCAQICEAMBmgI6/1ixa9bV6uYX8GYwDQYJKoZIhvcNAQEF
# BQAwYjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UE
# CxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJlZCBJ
# RCBDQS0xMB4XDTE0MTAyMjAwMDAwMFoXDTI0MTAyMjAwMDAwMFowRzELMAkGA1UE
# BhMCVVMxETAPBgNVBAoTCERpZ2lDZXJ0MSUwIwYDVQQDExxEaWdpQ2VydCBUaW1l
# c3RhbXAgUmVzcG9uZGVyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA
# o2Rd/Hyz4II14OD2xirmSXU7zG7gU6mfH2RZ5nxrf2uMnVX4kuOe1VpjWwJJUNmD
# zm9m7t3LhelfpfnUh3SIRDsZyeX1kZ/GFDmsJOqoSyyRicxeKPRktlC39RKzc5YK
# Z6O+YZ+u8/0SeHUOplsU/UUjjoZEVX0YhgWMVYd5SEb3yg6Np95OX+Koti1ZAmGI
# YXIYaLm4fO7m5zQvMXeBMB+7NgGN7yfj95rwTDFkjePr+hmHqH7P7IwMNlt6wXq4
# eMfJBi5GEMiN6ARg27xzdPpO2P6qQPGyznBGg+naQKFZOtkVCVeZVjCT88lhzNAI
# zGvsYkKRrALA76TwiRGPdwIDAQABo4IDNTCCAzEwDgYDVR0PAQH/BAQDAgeAMAwG
# A1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwggG/BgNVHSAEggG2
# MIIBsjCCAaEGCWCGSAGG/WwHATCCAZIwKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3
# LmRpZ2ljZXJ0LmNvbS9DUFMwggFkBggrBgEFBQcCAjCCAVYeggFSAEEAbgB5ACAA
# dQBzAGUAIABvAGYAIAB0AGgAaQBzACAAQwBlAHIAdABpAGYAaQBjAGEAdABlACAA
# YwBvAG4AcwB0AGkAdAB1AHQAZQBzACAAYQBjAGMAZQBwAHQAYQBuAGMAZQAgAG8A
# ZgAgAHQAaABlACAARABpAGcAaQBDAGUAcgB0ACAAQwBQAC8AQwBQAFMAIABhAG4A
# ZAAgAHQAaABlACAAUgBlAGwAeQBpAG4AZwAgAFAAYQByAHQAeQAgAEEAZwByAGUA
# ZQBtAGUAbgB0ACAAdwBoAGkAYwBoACAAbABpAG0AaQB0ACAAbABpAGEAYgBpAGwA
# aQB0AHkAIABhAG4AZAAgAGEAcgBlACAAaQBuAGMAbwByAHAAbwByAGEAdABlAGQA
# IABoAGUAcgBlAGkAbgAgAGIAeQAgAHIAZQBmAGUAcgBlAG4AYwBlAC4wCwYJYIZI
# AYb9bAMVMB8GA1UdIwQYMBaAFBUAEisTmLKZB+0e36K+Vw0rZwLNMB0GA1UdDgQW
# BBRhWk0ktkkynUoqeRqDS/QeicHKfTB9BgNVHR8EdjB0MDigNqA0hjJodHRwOi8v
# Y3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNybDA4oDag
# NIYyaHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ0Et
# MS5jcmwwdwYIKwYBBQUHAQEEazBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5k
# aWdpY2VydC5jb20wQQYIKwYBBQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydEFzc3VyZWRJRENBLTEuY3J0MA0GCSqGSIb3DQEBBQUAA4IB
# AQCdJX4bM02yJoFcm4bOIyAPgIfliP//sdRqLDHtOhcZcRfNqRu8WhY5AJ3jbITk
# WkD73gYBjDf6m7GdJH7+IKRXrVu3mrBgJuppVyFdNC8fcbCDlBkFazWQEKB7l8f2
# P+fiEUGmvWLZ8Cc9OB0obzpSCfDscGLTYkuw4HOmksDTjjHYL+NtFxMG7uQDthSr
# 849Dp3GdId0UyhVdkkHa+Q+B0Zl0DSbEDn8btfWg8cZ3BigV6diT5VUW8LsKqxzb
# XEgnZsijiwoc5ZXarsQuWaBh3drzbaJh6YoLbewSGL33VVRAA5Ira8JRwgpIr7DU
# buD0FAo6G+OPPcqvao173NhEMIIGzTCCBbWgAwIBAgIQBv35A5YDreoACus/J7u6
# GzANBgkqhkiG9w0BAQUFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNl
# cnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdp
# Q2VydCBBc3N1cmVkIElEIFJvb3QgQ0EwHhcNMDYxMTEwMDAwMDAwWhcNMjExMTEw
# MDAwMDAwWjBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkw
# FwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1
# cmVkIElEIENBLTEwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDogi2Z
# +crCQpWlgHNAcNKeVlRcqcTSQQaPyTP8TUWRXIGf7Syc+BZZ3561JBXCmLm0d0nc
# icQK2q/LXmvtrbBxMevPOkAMRk2T7It6NggDqww0/hhJgv7HxzFIgHweog+SDlDJ
# xofrNj/YMMP/pvf7os1vcyP+rFYFkPAyIRaJxnCI+QWXfaPHQ90C6Ds97bFBo+0/
# vtuVSMTuHrPyvAwrmdDGXRJCgeGDboJzPyZLFJCuWWYKxI2+0s4Grq2Eb0iEm09A
# ufFM8q+Y+/bOQF1c9qjxL6/siSLyaxhlscFzrdfx2M8eCnRcQrhofrfVdwonVnwP
# YqQ/MhRglf0HBKIJAgMBAAGjggN6MIIDdjAOBgNVHQ8BAf8EBAMCAYYwOwYDVR0l
# BDQwMgYIKwYBBQUHAwEGCCsGAQUFBwMCBggrBgEFBQcDAwYIKwYBBQUHAwQGCCsG
# AQUFBwMIMIIB0gYDVR0gBIIByTCCAcUwggG0BgpghkgBhv1sAAEEMIIBpDA6Bggr
# BgEFBQcCARYuaHR0cDovL3d3dy5kaWdpY2VydC5jb20vc3NsLWNwcy1yZXBvc2l0
# b3J5Lmh0bTCCAWQGCCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8A
# ZgAgAHQAaABpAHMAIABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQA
# aQB0AHUAdABlAHMAIABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUA
# IABEAGkAZwBpAEMAZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUA
# IABSAGUAbAB5AGkAbgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQA
# IAB3AGgAaQBjAGgAIABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEA
# bgBkACAAYQByAGUAIABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUA
# aQBuACAAYgB5ACAAcgBlAGYAZQByAGUAbgBjAGUALjALBglghkgBhv1sAxUwEgYD
# VR0TAQH/BAgwBgEB/wIBADB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0
# dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2NhY2Vy
# dHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDCBgQYD
# VR0fBHoweDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# QXNzdXJlZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQu
# Y29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDAdBgNVHQ4EFgQUFQASKxOY
# spkH7R7for5XDStnAs0wHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8w
# DQYJKoZIhvcNAQEFBQADggEBAEZQPsm3KCSnOB22WymvUs9S6TFHq1Zce9UNC0Gz
# 7+x1H3Q48rJcYaKclcNQ5IK5I9G6OoZyrTh4rHVdFxc0ckeFlFbR67s2hHfMJKXz
# BBlVqefj56tizfuLLZDCwNK1lL1eT7EF0g49GqkUW6aGMWKoqDPkmzmnxPXOHXh2
# lCVz5Cqrz5x2S+1fwksW5EtwTACJHvzFebxMElf+X+EevAJdqP77BzhPDcZdkbkP
# Z0XN1oPt55INjbFpjE/7WeAjD9KqrgB87pxCDs+R1ye3Fu4Pw718CqDuLAhVhSK4
# 6xgaTfwqIa1JMYNHlXdx3LEbS0scEJx3FMGdTy9alQgpECYxggQ7MIIENwIBATCB
# hjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQL
# ExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3Vy
# ZWQgSUQgQ29kZSBTaWduaW5nIENBAhAG8IyaFm1PlWkZ0FW8vvwPMAkGBSsOAwIa
# BQCgeDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgor
# BgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3
# DQEJBDEWBBR+/zDDckWtwxRHDQ2p4T0/bBnr7TANBgkqhkiG9w0BAQEFAASCAQCL
# F0j0YrImVl+3QxdxLFFqire39ERD+gt+lhKqdenQkDpE2qN/W7/rXi3Nlj/jpp+L
# pM3hg+wPgS/cC85I4k5TCuHeLdPlC6XzT3tBFWijM5omr2WlrvgeZBvGxQiC6F3d
# t9AAPmnw+aSEtF99AMbTGi3PgPEHNZDrjKL8Iq4OOBQgTiSqp2rSfQjiMXd0yd4U
# ZmJJy+mWdNxVATjg5fbPG53wM71+GjxXySndd6sZLZDv4H3xU1gCQFMigqRS1/de
# 6uHBaHYBXE9iSVU6LUFGlSZvAu5igV7wcPLnGPRnbVh7L7b/2HO2+93b/12L5yJj
# TcUEppk9Xp1dn0Sj+KQroYICDzCCAgsGCSqGSIb3DQEJBjGCAfwwggH4AgEBMHYw
# YjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQ
# d3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJlZCBJRCBD
# QS0xAhADAZoCOv9YsWvW1ermF/BmMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMx
# CwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xODA5MTMwMjA2NTlaMCMGCSqG
# SIb3DQEJBDEWBBRYcWWcXLkz5srKsJagwzZBpObyQzANBgkqhkiG9w0BAQEFAASC
# AQAT5uQKaCcVdExOz0N2s3ZAjoPb8OLOaFh3U/6A6E5MNOpvmo2NF51xsaav6yUD
# B40ZkqbI1ewoYDVZ5YMwmt7AfpU7ejNpLaA8QuEzaPOYYSupsTy2de59v0RiGZWE
# JFzX5aJj+wf/qlLooK29zW3S79LhzRj5Weve9/YfjyIb3pfjoZ3zatMdvRWxI8h6
# UYpCWXGm21QIj36kPXlO5tu1pNB3teWc7CfPZuvUtuBWnXXXLStBIAEwMwVUZ2cF
# QwWXlEoWblVQfQb9wCAGHmD/Wvy4KDVtFJY7GrAiiVCAFwhUnaIdKo2P3INK+cfc
# Z3Rv7Q5DY0a4j1v3DoWYlwbg
# SIG # End signature block
