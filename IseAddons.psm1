Function Get-IseAddonsCaretValue
{
    $result = 1 | Select-Object Title, Code, Value
    $result.Title = 'Value at Caret'

    Write-Host -ForegroundColor Cyan ("="*100)
    Write-Host -ForegroundColor Cyan "Executing ISE Caret/Selection Value"
    Write-Host -ForegroundColor Cyan ("="*100)
    Write-Host ""

    # If text is currently highlighted, execute this code
    $selectedText = $psISE.CurrentFile.Editor.SelectedText
    if ($selectedText)
    {
        $result.Code = $selectedText
        $result.Title = $selectedText -split "[`r`n]+" |
            Select-Object -Last 10 |
            ForEach-Object Trim
        $result.Title = $result.Title -join ' '
        Write-Host -ForegroundColor Cyan $result.Code
        Write-Host ""
        $result.Value = & ([ScriptBlock]::Create($selectedText))
        return $result
    }

    # Try to find out if the caret is currently on a variable name
    $column = $psISE.CurrentFile.Editor.CaretColumn
    $parseErrors = $null
    $tokenList = [System.Management.Automation.PSParser]::Tokenize($psISE.CurrentFile.Editor.CaretLineText, [ref]$parseErrors)
    $caretToken = $tokenList |
        Where-Object StartColumn -le $column |
        Where-Object EndColumn -ge $column |
        Select-Object -First 1
    if ($caretToken.Type -eq 'Variable')
    {
        # Return the value of the variable if the caret is on a variable name
        $result.Code = "`$$($caretToken.Content)"
        Write-Host -ForegroundColor Cyan $result.Code
        Write-Host ""
        $result.Value = Get-Variable -Name $caretToken.Content -ValueOnly |
            ForEach-Object { $_ } # This will force expansion of an array
        $result.Title = "Value of `$$($caretToken.Content)"
        return $result
    }
    
    # Otherwise work backwards to find the beginning of the current pipeline
    $fileLineList = $psISE.CurrentFile.Editor.Text -split "`r`n"
    $lineNumber = $psISE.CurrentFile.Editor.CaretLine

    $codeLineList = New-Object System.Collections.Generic.List[string]
    $thisLine = $fileLineList[$lineNumber-1].TrimEnd().TrimEnd("|").TrimEnd()
    $lineNumber -= 1
    $codeLineList.Add($thisLine)

    # Keep track of how deep in a scriptblock we are to stay at the same depth
    $blockUp = $thisLine.Length - $thisLine.Replace('}', '').Length
    $blockDown = $thisLine.Length - $thisLine.Replace('{', '').Length
    $blockCount = $blockUp - $blockDown

    # Keep track of the type of Here-String we're in
    $inHereString = if ($thisLine.StartsWith("'@")) { "'" }
    elseif ($thisLine.StartsWith('"@')) { '"' }

    # Loop backwards as long as we still have lines to go until we appear to be at the beginning of the statement.
    # We'll know we're at the beginning when the PREVIOUS line in the file doesn't end with a pipe, comma, backtick,
    # or open curly bracket, and we're not currently inside a Here-String or a ScriptBlock.
    while ($lineNumber -gt 0)
    {
        $thisLine = $fileLineList[$lineNumber-1].TrimEnd()
        if (!$inHereString)
        {
            $inHereString = if ($thisLine.StartsWith("'@")) { "'" }
            elseif ($thisLine.StartsWith('"@')) { '"' }
        }
        if ($inHereString)
        {
            if ($thisLine.Contains("@$inHereString")) { $inHereString = $false }
            $codeLineList.Add($thisLine)
            $lineNumber -= 1
            continue
        }
        $continue = $thisLine.EndsWith("|") -or $thisLine.EndsWith('`') -or $thisLine.EndsWith(',') -or $thisLine.EndsWith('{')
        # Skip over lines with comments
        if ($thisLine.Trim().StartsWith('#'))
        {
            $lineNumber -= 1
            continue
        }
        $blockUp = $thisLine.Length - $thisLine.Replace('}', '').Length
        $blockDown = $thisLine.Length - $thisLine.Replace('{', '').Length
        $blockCount = $blockCount + $blockUp - $blockDown
        if ((!$continue -and $blockCount -le 0) -or $blockCount -lt 0) { break }
        $codeLineList.Add($thisLine)
        $lineNumber -= 1
    }

    # Put the code back in order
    $codeLineList.Reverse()

    $result.Title = $codeLineList | Select-Object -Last 10 | ForEach-Object Trim
    $result.Title = $result.Title -join ' '
    $result.Code = $codeLineList -join "`r`n" -replace "[ \t\|]+\Z" -replace "(?m)\A *\$.+= ?"

    Write-Host -ForegroundColor Cyan $result.Code
    Write-Host ""
    $result.Value = & ([ScriptBlock]::Create($result.Code))
    return $result
}

& {
    trap { continue } # Easier to just ignore all exceptions here

    $psISE.CurrentPowerShellTab.AddOnsMenu.Submenus.Add("Get Caret/Selected Value (First 1)",
        { Get-IseAddonsCaretValue | ForEach-Object Value | Select-Object -First 1 | Format-List -Property * },
        "Alt+F2"
    )

    $psISE.CurrentPowerShellTab.AddOnsMenu.Submenus.Add("Get Caret/Selected Value (Out GridView)",
        { Get-IseAddonsCaretValue | ForEach-Object Value | Out-GridView },
        "F2"
    )

    $psISE.CurrentPowerShellTab.AddOnsMenu.Submenus.Add("Get Caret/Selected Value (Out GridView PassThru)",
        { Get-IseAddonsCaretValue | ForEach-Object Value | Out-GridView -PassThru },
        "Shift+F2"
    )

    $psISE.CurrentPowerShellTab.AddOnsMenu.Submenus.Add("Get Caret/Selected Value (Count)",
        { Get-IseAddonsCaretValue | ForEach-Object Value | Measure-Object | ForEach-Object Count },
        "Alt+Shift+F2"
    )

}
