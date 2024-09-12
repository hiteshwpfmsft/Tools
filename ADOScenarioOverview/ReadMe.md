# ADO Scenario Report Generator

## Overview

This project generates comprehensive Scenario reports for contributors and deliverables, providing summaries and detailed information in a well-organized format. It uses the `ClosedXML` library to work with Excel files.

## Features

- **Scenario Summary**: Provides an overview of contributor capacities and commitments.
- **Sprint Overview**: Aggregates and presents sprint data by contributor.
- **Detailed Deliverable Sheets**: Lists tasks and contributors with estimated and remaining days.
- **Deliverable Details**: Offers a detailed view of deliverables, tasks, and their statuses.

## Prerequisites

- .NET 6.0 or later
- [ClosedXML](https://github.com/ClosedXML/ClosedXML) library

## Installation

1. **Clone the Repository**
    ```bash
    git clone <repository_url>
    cd <repository_directory>
    ```

2. **Install Dependencies**
    ```bash
    dotnet add package ClosedXML
    ```

3. **Build the Project**
    ```bash
    dotnet build
    or build the project from VS
    ```

## Configuration

The project uses a `Config` class for configuration. Update the following parameters in `Config.cs` to suit your needs:

- `PersonalAccessToken`: Token for authenticating API requests.
- `ScenarioID`: ID for the specific scenario to generate reports.
- `SpecialContributors`: Set of contributors with special capacity multipliers.
- `CapacityMultiplierSpecial`: Multiplier for special contributors.
- `CapacityMultiplierRegular`: Multiplier for regular contributors.
- `EndDate`: End date for the reporting period.
- `CurrentDate`: The current date for generating the report.
- `Organization`: Azure DevOps organization name.
- `Project`: Azure DevOps project name.
- `ApiVersion`: API version for Azure DevOps requests.
- `TasksBaseUri`: Base URI for Azure DevOps task API.
- `TasksBaseUriEdit`: Base URI for editing tasks in Azure DevOps.
- `ExcelFilePath`: Path and name template for the generated Excel file.

