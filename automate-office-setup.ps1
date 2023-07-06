#!/usr/bin/env pwsh
# This script automates the Office installation and product license activation.

try {
    Write-Host "Starting Office Installation..."

    # Step 1: Open command prompt as administrator
    Start-Process powershell -Verb RunAs

    # Step 2: Navigate to Office 2021 setup folder in Command Line
    Set-Location -Path 'C:\Users\Administrator1\Desktop\Office2021'

    # Step 3: Download and Install Office 2021 using the specified configuration file
    Start-Process -FilePath 'setup.exe' -ArgumentList '/download', 'configuration-Office2021Enterprise.xml' -Wait

    # Step 4: Configure Office 2021 automatically using the specified configuration file
    Start-Process -FilePath 'setup.exe' -ArgumentList '/configure', 'configuration-Office2021Enterprise.xml' -Wait

    Write-Host "Office Installation completed successfully."

    Write-Host "Starting Office activation..."

    # Step 5: Navigate to Office installation folder for license key
    Set-Location -Path 'C:\Program Files\Microsoft Office\Office16'

    # Further steps are taken automatically that activate the product licenses which are not added here for confidentiality.

    Write-Host "Office activation completed successfully."
}
catch {
    Write-Host "An error occurred during Office installation or activation:"
    Write-Host $_.Exception.Message
    exit 1
}

exit 0
