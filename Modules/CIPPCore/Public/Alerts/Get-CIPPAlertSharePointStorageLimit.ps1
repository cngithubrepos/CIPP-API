function Get-CIPPAlertSharePointStorageLimit {
    <#
    .SYNOPSIS
        Alerts when a tenant's SharePoint storage usage exceeds a defined threshold.
    .DESCRIPTION
        Queries the Microsoft Graph Reports API for SharePoint storage usage and
        fires an alert if the tenant's storage consumption exceeds the threshold
        percentage set below.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)]
        [Alias('Tenant')]
        [System.String]$TenantFilter
    )

    # --- Set your alert threshold here (percentage) ---
    $ThresholdPercent = 90

    Write-LogMessage -API 'Alerts' -tenant $TenantFilter -message 'Checking SharePoint tenant storage usage' -sev Debug

    try {
        # Graph Reports API: SharePoint storage usage over the last 7 days
        # Returns a CSV with columns: Report Date, Site Type, Storage Used (Byte),
        # Storage Allocated (Byte), Report Period
        $ReportData = New-GraphGetRequest `
            -uri "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageStorage(period='D7')" `
            -tenantid $TenantFilter `
            -AsApp $true

        # Parse CSV and get the most recent 'All' row
        $LatestRow = $ReportData |
            ConvertFrom-Csv |
            Where-Object { $_.'Site Type' -eq 'All' } |
            Sort-Object 'Report Date' -Descending |
            Select-Object -First 1

        if ($LatestRow) {
            $UsedBytes  = [long]$LatestRow.'Storage Used (Byte)'
            $AllocBytes = [long]$LatestRow.'Storage Allocated (Byte)'

            if ($AllocBytes -gt 0) {
                $PercentUsed = [math]::Round(($UsedBytes / $AllocBytes) * 100, 1)
                $UsedGB      = [math]::Round($UsedBytes  / 1GB, 2)
                $TotalGB     = [math]::Round($AllocBytes / 1GB, 2)

                if ($PercentUsed -ge $ThresholdPercent) {
                    $AlertData = @(
                        [PSCustomObject]@{
                            Tenant      = $TenantFilter
                            PercentUsed = "$PercentUsed%"
                            'Used (GB)' = $UsedGB
                            'Total (GB)'= $TotalGB
                            Message     = "SharePoint storage at $PercentUsed% — $UsedGB GB of $TotalGB GB used"
                        }
                    )

                    if ($AlertData) {
                        Write-AlertTrace -cmdletName $MyInvocation.MyCommand -tenantFilter $TenantFilter -data $AlertData
                    }
                }
            }
        }
    } catch {
        Write-LogMessage -API 'Alerts' -tenant $TenantFilter `
            -message "Could not get SharePoint storage for $($TenantFilter): $(Get-NormalizedError -message $_.Exception.message)" `
            -sev Error
    }
}
