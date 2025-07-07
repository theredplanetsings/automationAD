<#============================================
  Export All AD Groups Script
  
  This script searches the entire paradigmcos.local domain
  for all distribution and security groups and exports
  them to a text file with one group per line.
============================================#>

# =========================
# Imports and Initial Setup
# =========================
Import-Module ActiveDirectory

# =========================
# Configuration
# =========================
$DomainDN = "DC=paradigmcos,DC=local"
$OutputFile = "AD_Groups_Export.txt"
$OutputPath = Join-Path $PSScriptRoot $OutputFile
# =========================
# Function to Export Groups
# =========================
function Export-ADGroups {
    param(
        [string]$Domain,
        [string]$OutputFilePath
    )
    
    Write-Host "Starting AD Groups export..." -ForegroundColor Green
    Write-Host "Domain: $Domain" -ForegroundColor Cyan
    Write-Host "Output File: $OutputFilePath" -ForegroundColor Cyan
    
    try {
        # Search for all groups in the domain
        Write-Host "Searching for all distribution and security groups..." -ForegroundColor Yellow
        
        $AllGroups = Get-ADGroup -Filter * -SearchBase $Domain -Properties Name, GroupCategory, GroupScope, Description, DistinguishedName | Sort-Object Name
        
        if ($AllGroups) {
            Write-Host "Found $($AllGroups.Count) groups. Writing to file..." -ForegroundColor Yellow
            
            # Create header for the file
            $Header = @"
Active Directory Groups Export
Domain: $Domain
Export Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Total Groups Found: $($AllGroups.Count)

Format: GroupName | GroupType | GroupScope | Description | Distinguished Name
================================================================================

"@
            
            # Write header to file
            $Header | Out-File -FilePath $OutputFilePath -Encoding UTF8
            
            # Write each group to file
            foreach ($Group in $AllGroups) {
                $GroupType = $Group.GroupCategory
                $GroupScope = $Group.GroupScope
                $Description = if ($Group.Description) { $Group.Description } else { "No description" }
                $DN = $Group.DistinguishedName
                
                $GroupLine = "$($Group.Name) | $GroupType | $GroupScope | $Description | $DN"
                $GroupLine | Out-File -FilePath $OutputFilePath -Append -Encoding UTF8
            }
            
            # Create summary statistics
            $SecurityGroups = ($AllGroups | Where-Object { $_.GroupCategory -eq "Security" }).Count
            $DistributionGroups = ($AllGroups | Where-Object { $_.GroupCategory -eq "Distribution" }).Count
            
            $Summary = @"

================================================================================
SUMMARY STATISTICS
================================================================================
Total Groups: $($AllGroups.Count)
Security Groups: $SecurityGroups
Distribution Groups: $DistributionGroups

Group Scope Breakdown:
"@
            
            $Summary | Out-File -FilePath $OutputFilePath -Append -Encoding UTF8
            
            # Add scope breakdown
            $ScopeBreakdown = $AllGroups | Group-Object GroupScope | ForEach-Object {
                "  $($_.Name): $($_.Count)"
            }
            
            $ScopeBreakdown | Out-File -FilePath $OutputFilePath -Append -Encoding UTF8
            
            Write-Host "Export completed successfully!" -ForegroundColor Green
            Write-Host "File saved to: $OutputFilePath" -ForegroundColor Green
            Write-Host "Total groups exported: $($AllGroups.Count)" -ForegroundColor Green
            Write-Host "  - Security Groups: $SecurityGroups" -ForegroundColor Cyan
            Write-Host "  - Distribution Groups: $DistributionGroups" -ForegroundColor Cyan
            
        } else {
            Write-Host "No groups found in the domain." -ForegroundColor Red
            "No groups found in domain: $Domain`nExport Date: $(Get-Date)" | Out-File -FilePath $OutputFilePath -Encoding UTF8
        }
        
    } catch {
        $ErrorMessage = "Error exporting groups: $($_.Exception.Message)"
        Write-Host $ErrorMessage -ForegroundColor Red
        $ErrorMessage | Out-File -FilePath $OutputFilePath -Encoding UTF8
        throw
    }
}

# =========================
# Function to Create Simple List
# =========================
function Export-SimpleGroupList {
    param(
        [string]$Domain,
        [string]$OutputFilePath
    )
    
    $SimpleOutputFile = $OutputFilePath -replace '\.txt$', '_Simple.txt'
    
    try {
        Write-Host "Creating simple group list..." -ForegroundColor Yellow
        
        $AllGroups = Get-ADGroup -Filter * -SearchBase $Domain | Sort-Object Name
        
        if ($AllGroups) {
            # Create simple list with just group names
            "# Simple AD Groups List - One group per line" | Out-File -FilePath $SimpleOutputFile -Encoding UTF8
            "# Domain: $Domain" | Out-File -FilePath $SimpleOutputFile -Append -Encoding UTF8
            "# Export Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File -FilePath $SimpleOutputFile -Append -Encoding UTF8
            "# Total Groups: $($AllGroups.Count)" | Out-File -FilePath $SimpleOutputFile -Append -Encoding UTF8
            "" | Out-File -FilePath $SimpleOutputFile -Append -Encoding UTF8
            
            foreach ($Group in $AllGroups) {
                $Group.Name | Out-File -FilePath $SimpleOutputFile -Append -Encoding UTF8
            }
            
            Write-Host "Simple list created: $SimpleOutputFile" -ForegroundColor Green
        }
        
    } catch {
        Write-Host "Error creating simple list: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# =========================
# Main Execution
# =========================
try {
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host "  Active Directory Groups Export Tool  " -ForegroundColor Magenta
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host ""
    
    # Check if AD module is available
    if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
        throw "Active Directory module is not available. Please install RSAT tools."
    }
    
    # Test AD connectivity
    Write-Host "Testing Active Directory connectivity..." -ForegroundColor Yellow
    try {
        $null = Get-ADDomain -Identity "paradigmcos.local"
        Write-Host "✓ Successfully connected to paradigmcos.local domain" -ForegroundColor Green
    } catch {
        throw "Unable to connect to paradigmcos.local domain. Please check your connection and permissions."
    }
    
    # Export detailed group information
    Export-ADGroups -Domain $DomainDN -OutputFilePath $OutputPath
    
    # Export simple group list
    Export-SimpleGroupList -Domain $DomainDN -OutputFilePath $OutputPath
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host "Export completed successfully!" -ForegroundColor Green
    Write-Host "Files created:" -ForegroundColor Green
    Write-Host "  1. $OutputPath (detailed)" -ForegroundColor Cyan
    Write-Host "  2. $($OutputPath -replace '\.txt$', '_Simple.txt') (names only)" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Magenta
    
} catch {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    exit 1
}

# =========================
# Optional: Open output folder
# =========================
$OpenFolder = Read-Host "`nWould you like to open the output folder? (Y/N)"
if ($OpenFolder -match '^[Yy]') {
    try {
        Invoke-Item $PSScriptRoot
    } catch {
        Write-Host "Could not open folder automatically. Files are located at: $PSScriptRoot" -ForegroundColor Yellow
    }
}
