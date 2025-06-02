# region Set- Cmdlet
Function Set-WindowTitle {
    <#
        .SYNOPSIS
            This cmdlet will set the window title for the current PowerShell session

        .PARAMETER Title
            This is a mandatory parameter which specifies the window title name.

        .INPUTS
            [string]

        .OUTPUTS
            None.

        .EXAMPLE
            Set-WindowTitle "Client Azure AD"

            The preceding example set the title of the current PowerShell window to be "Client Azure AD"

        .NOTES
            Version:
                - 5.1.2023.0712:    New function
    #>
    [CmdletBinding()]
    [OutputType([System.Void])]
    Param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, 
            HelpMessage = "Specify the desired PowerShell Window Title")]
        [string]$Title
    )

    begin { 
    }#begin
    
    process {
        if ($Title -ceq "ECTO 1") {
            $host.ui.RawUI.WindowTitle = "Listen...you smell something?"
        } else {
            $host.ui.RawUI.WindowTitle = $Title
        }
    }#process

    end {
    }#end 
}#Function Set-WindowTitle

#endregion

#region Export Module Members

Export-ModuleMember -Function Set-WindowTitle

#endregion