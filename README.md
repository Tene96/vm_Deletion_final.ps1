# VMware VM Deletion Automation

PowerShell automation tool designed to automate and standardize the VM decommissioning process across multiple VMware environments.

Instead of manually removing a VM from each infrastructure system, the script performs the decommissioning workflow from a single interface.

## What the Script Does

The script automates the VM decommissioning process across:

- VMware vCenter
- Active Directory
- SCCM
- DNS
- SQL
- Email notifications

It identifies the VM and its environment, performs validation checks, removes the relevant infrastructure objects, records the operation in SQL, and sends a completion notification.

## Multi-Environment Support

The script is designed to work with multiple VMware environments.

vCenter environments are configured in `config.json`, allowing additional environments to be added without creating separate scripts for each vCenter.

## Requirements

- PowerShell
- VMware PowerCLI
- VMware vCenter access
- Active Directory PowerShell module
- SCCM access
- DNS permissions
- Microsoft SQL Server
- SMTP access

## Configuration

Before running the script, edit `config.json` and configure:

- vCenter servers
- SCCM Server and Site Code
- DNS Server and Zone
- Active Directory settings
- SQL Server and Database
- SQL credentials
- SMTP / Email settings

Also review the environment-specific arrays inside `vm_Deletion_final.ps1` and adjust them to match your infrastructure.

## Important

VM deletion is a destructive operation.

Always validate the configuration and test the script in a non-production environment before using it in production.
