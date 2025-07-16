# Intune App Uploader

A PowerShell GUI tool for automating Winget app deployments through Microsoft Intune. Note that Windows Defender could block the script on first run. Unblock the script in Windows Defender to solve this problem.

## Overview

This script provides a graphical interface to:
<ul>
    <li>Register and consent an Azure AD application for Intune management</li>
    <li>Search and select Winget packages</li>
    <li>Package and upload applications to Intune</li>
    <li>Automatically create detection and installation scripts</li>
    <li>Deploy applications to all devices, all users and set required or available</li>
</ul>

## Prerequisites

<ul>
    <li>Windows 10/11</li>
    <li>PowerShell 7</li>
    <li>Admin rights on the local machine</li>
    <li>Intune admin access</li>
    <li>Internet connectivity</li>
</ul>

## Setup Process

### First-Time Setup

<ol>
    <li>Run the script as administrator.</li>
    <li>Click the "App Reg" button to:
        <ul>
            <li>Register an Azure AD application</li>
            <li>Create necessary permissions</li>
            <li>Generate client credentials</li>
            <li>Automatically grant admin consent to the app</li>
        </ul>
    </li>
</ol>

## Adding Applications

You can add applications in two ways:

### Via Winget Search

<ol>
    <li>Click "Search Winget".</li>
    <li>Enter a search term.</li>
    <li>Select an application from the results.</li>
    <li>Add publisher and description details.</li>
    <li>Click OK to save.</li>
</ol>

### Manual Addition

<ol>
    <li>Click "Add".</li>
    <li>Fill in:
        <ul>
            <li>Display Name</li>
            <li>Winget ID</li>
            <li>Publisher</li>
            <li>Description</li>
        </ul>
    </li>
    <li>Click OK to save.</li>
</ol>

## Deploying Applications

<ol>
    <li>Select one or more applications from the list.
        <ul>
            <li>If none are selected, you'll be prompted to process all apps.</li>
        </ul>
    </li>
    <li>Click "Run" to start deployment.
        <ul>
            <li>Choose if you want to assign the "All Devices" or "All Users" group as required or available.</li>
        </ul>
    </li>
    <li>The script will:
        <ul>
            <li>Create installation scripts</li>
            <li>Create uninstallation scripts</li>
            <li>Create detection scripts</li>
            <li>Package everything using IntuneWinAppUtil</li>
            <li>Upload to Intune</li>
            <li>Create device assignments</li> 
        </ul>
    </li>
</ol>

## Additional Features

<ul>
    <li><strong>Edit</strong>: Modify existing application details.</li>
    <li><strong>Remove</strong>: Delete applications from the list.</li>
    <li><strong>Del Credentials</strong>: Remove stored Azure AD credentials.</li>
    <li><strong>Logging</strong>: Detailed logs stored in the Logs directory.</li>
</ul>

## Configuration

<p>Settings are stored in: <code>[Script Directory]/config.json</code></p>
<p>Includes:</p>
<ul>
    <li>Application list</li>
    <li>Azure AD credentials</li>
    <li>Application configurations</li>
</ul>

## Cleanup

<p>The script automatically:</p>
<ul>
    <li>Removes temporary files after deployment</li>
    <li>Cleans up packaging directories</li>
    <li>Maintains organized log files</li>
    <li>Deletes credentials</li>
    <li>Deletes App registration</li>
</ul>

## Logging

<p>All actions are logged to: <code>[Script Directory]/Logs/IntuneUploadLog.txt</code></p>
<p>Individual app installation/uninstallation logs are created during deployment and can be found at <code>C:\ProgramData\Microsoft\IntuneManagementExtension\Logs</code>.</p>
