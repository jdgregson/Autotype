Param (
    [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true
    )]
    [String]$Text,

    [Alias("Timeout")]
    [Int]$Delay = 5,

    [Switch]$Test
)

function Do-Countdown {
    [CmdletBinding()]
    Param (
        [Parameter(
            Mandatory = $true
        )]
        [Int]$Delay
    )

    begin {
        $title = "CHANGE TO TARGET WINDOW"
        $text = "Autotyping text in"

        Write-Host ("`n" * 5)
        $host.ui.RawUI.WindowTitle = "$title - Autotype"
    }
    process {
        for ($i = 0; $i -lt $Delay; $i++) {
            Write-Progress -Activity $title -Status "$text $($Delay - $i) seconds..." -PercentComplete ($i/$Delay*100)
            Write-Warning $title
            Start-Sleep 1
        }
    }
}

function Do-Autotype {
    [CmdletBinding()]
    Param (
        [Parameter(
            Mandatory = $true
        )]
        [String]$Text
    )

    begin {
        $wshell = New-Object -ComObject wscript.shell
        $Text = $Text -replace "\{", "@@@{@@@"
        $Text = $Text -replace "\}", "@@@}@@@"
        $Text = $Text -replace "@@@{@@@", "{{}"
        $Text = $Text -replace "@@@}@@@", "{}}"
        $Text = $Text -replace "\(", "{(}"
        $Text = $Text -replace "\)", "{)}"
        $Text = $Text -replace "\!", "{!}"
        $Text = $Text -replace "\^", "{^}"
        $Text = $Text -replace "\+", "{+}"
        $Text = $Text -replace "%", "{%}"
        $Text = $Text -replace "~", "{~}"
    }
    process {
        if ($Test) {
            $Text
        }
        $wshell.SendKeys($Text)
    }
}

function Do-Test() {
    begin {
        $wshell = New-Object -ComObject wscript.shell
        $testText = "abcdefghijklmnopqrstuvwxyz" +
            "ABCDEFGHIJKLMNOPQRSTUVWXYZ" +
            "``1234567890-=[]\;',./~!@#$%^&*()_+{}|:`"<>?"
        $file = $env:TEMP + "\autotype-text.txt"
    }
    process {
        Set-Content -Path $file -Value "" -NoNewline
        notepad.exe $file
        Start-Sleep -Seconds 1

        Do-Autotype -Text $testText
        $wshell.SendKeys("^(s)")
        $wshell.SendKeys("%({F4})")

        $writtenText = Get-Content $file -Raw
        if ($writtenText -and $writtenText -eq $testText) {
            Write-Host "Test passed"
        } else {
            Write-Warning "Test failed"
        }
    }
}

if ($Test) {
    Do-Test
} else {
    Do-Countdown -Delay $Delay
    Do-Autotype -Text $Text
}
