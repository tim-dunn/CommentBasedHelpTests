# CommentBasedHelp.test.ps1
Pester `.test.ps1` file for PSM1 functions' comment based help.

## Criteria
All exported functions have `Get-Help` output with:
+ `.SYNOPSIS`
+ `.DESCRIPTION`
+ `.INPUTS`
+ `.EXAMPLE`

All function parameters have:
- `.PARAMETER`
- Default value, _or_ `[Parameter( Mandatory )]` attribute set
- sample usage in at least one `.EXAMPLE`

## Usage
Either copy the `Describe` block into your existing `.tests.ps1` file, or copy the file into the PSModule's root folder. If using the latter method, trigger the tests by running `Invoke-Pester` from the PSModule's root folder or some ancestor folder.

```
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
```
