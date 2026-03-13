# O365 License Cost Optimizer

## Overview
This PowerShell automation script identifies inactive Office 365 user accounts and generates automated reports. It was developed to reduce unnecessary licensing costs by detecting accounts dormant for over 3 months.

## Features
* Queries On-Premise Active Directory for user last logon timestamps.
* Cross-references AD data with active O365 licenses.
* Automatically dispatches a formatted email report to the IT administration team.
* Designed to run autonomously via Windows Task Scheduler.

## Tech Stack
* **Language:** PowerShell
* **Technologies:** Microsoft 365, On-Premise Active Directory

## Note
*This repository contains a sanitized version of the script. Hardcoded IPs, domains, and credentials have been removed for security purposes.*
