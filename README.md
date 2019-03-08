# IseAddons

Contains hotkeys for the PowerShell ISE to help with coding and testing. Simply importing this module will add the hotkeys (and the corresponding menu items).

# Caret/Selection Value Hotkeys

Option 1: Put your caret (the blinking cursor) on a line or a variable name. Press the F2 key. If on a variable name, the addon will retrieve the value of the variable and send it to Out-GridView. If not on a variable name, the addon will execute the entire line, working backwards in the file until it finds what it thinks in the beginning of the current pipeline.

Option 2: Highlight some text. Press the F2 key. The addon will execute the highlighted text as a scriptblock and send the result to Out-GridView.

## Hotkeys
* F2 - Out-GridView
* Alt+F2 - Select-Object -First 1 | Format-List -Property *
* Shift+F2 - Out-GridView -PassThru
* Alt+Shift+F2 - Measure-Object | ForEach-Object Count

## Detection
Take the following PowerShell code as a sample:

```powershell
Get-ChildItem C:\Windows |
    Select-Object Name, Length, LastWriteTime |
    Where-Object Length |
    ForEach-Object {
        $_.Length = "$([Math]::Round($_.Length / 1MB,2)) MB"
        $_
    }
```

Putting your caret (click anywhere on the line without highlighting any text) on the first line (Get-ChildItem) and pressing Alt+F2 will return the first item as though you had executed:
```powershell
Get-ChildItem | Select-Object -First 1 | Format-List -Property *
```
You'll also see the text "Get-ChildItem C:\Windows" in Cyan preceded by a header. Moving your caret to the second line by pressing the Down Arrow key and pressing Alt+F2 will show that you executed two lines. Using this addon (and assuming you write scripts with pipeline statements on separate lines) you should be able to walk through a pipeline and quickly preview the results at each step.

You'll notice if you put your caret inside of the ForEach-Object loop you'll only execute the one line in the loop. If you put your caret on } on the last line you'll execute everything. Just look at the "Executing ISE Caret/Selection Value" text to get a feel for how it works. Also remember that highlighting any text will execute just the highlighted text, and if your caret is on a variable name you'll get just the contents of the variable.

The addon will also ignore any "$variable = " text since it wants to show you the value of your selection, not capture it in a variable.
