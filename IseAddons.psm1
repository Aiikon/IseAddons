Function Get-IseAddonsCaretValue
{
    $private:result = 1 | Select-Object Title, Code, Value
    $result.Title = 'Value at Caret'

    Write-Host -ForegroundColor Cyan ("="*100)
    Write-Host -ForegroundColor Cyan "Executing ISE Caret/Selection Value"
    Write-Host -ForegroundColor Cyan ("="*100)
    Write-Host ""

    # If text is currently highlighted, execute this code
    $private:selectedText = $psISE.CurrentFile.Editor.SelectedText
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
    $private:column = $psISE.CurrentFile.Editor.CaretColumn
    $private:parseErrors = $null
    $private:tokenList = [System.Management.Automation.PSParser]::Tokenize($psISE.CurrentFile.Editor.CaretLineText, [ref]$parseErrors)
    $private:caretToken = $tokenList |
        Where-Object StartColumn -le $column |
        Where-Object EndColumn -ge $column |
        Select-Object -First 1
    if ($caretToken.Type -eq 'Variable')
    {
        # Return the value of the variable if the caret is on a variable name
        $result.Code = "`$$($caretToken.Content)"
        Write-Host -ForegroundColor Cyan $result.Code
        Write-Host ""
        $result.Value = & { Param($_) Get-Variable -Name $_ -ValueOnly } $caretToken.Content |
            ForEach-Object { $_ } # This will force expansion of an array
        $result.Title = "Value of `$$($caretToken.Content)"
        return $result
    }
    
    # Otherwise work backwards to find the beginning of the current pipeline
    $private:fileLineList = $psISE.CurrentFile.Editor.Text -split "`r`n"
    $private:lineNumber = $psISE.CurrentFile.Editor.CaretLine

    $private:codeLineList = New-Object System.Collections.Generic.List[string]
    $private:thisLine = $fileLineList[$lineNumber-1].TrimEnd().TrimEnd("|").TrimEnd()
    $lineNumber -= 1
    $codeLineList.Add($thisLine)

    # Loop backwards as long as we still have lines to go until we appear to be at the beginning of the statement.
    # We'll know we're at the beginning when the PREVIOUS line in the file doesn't end with a pipe, comma, backtick,
    # or open curly bracket, and the scriptblock is valid
    while ($lineNumber -gt 0)
    {
        $thisLine = $fileLineList[$lineNumber-1].TrimEnd()
        $continue = $thisLine.EndsWith("|") -or $thisLine.EndsWith('`') -or $thisLine.EndsWith(',') -or
            (!$(try { [ScriptBlock]::Create($codeLineList -join "`r`n") } catch {}))
        # Skip over lines with comments
        if ($thisLine.Trim().StartsWith('#'))
        {
            $lineNumber -= 1
            continue
        }
        if (!$continue) { break }
        $codeLineList.Insert(0, $fileLineList[$lineNumber-1])
        $lineNumber -= 1
    }

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
