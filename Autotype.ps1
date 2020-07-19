Param (
    [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true
    )]
    [String]$Text,

    [Alias("Timeout")]
    [Int]$Delay = 5
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
    }
    process {
        $wshell.SendKeys($Text)
    }
}

Do-Countdown -Delay $Delay
Do-Autotype -Text $Text
