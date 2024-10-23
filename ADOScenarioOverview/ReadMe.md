# ADO Scenario Hub

## Overview

The **ADO Scenario Hub** is a comprehensive solution designed to streamline the management of Azure DevOps scenarios, deliverables, and reports. This hub combines tools for creating, organizing, and reporting on work items, ensuring efficient collaboration and tracking within teams. Utilizing the `ClosedXML` library, it facilitates seamless interaction with Excel files for data input and reporting.

## Projects

### 1. ADO Scenario Planner

The **ADO Scenario Planner** tool automates the creation of Azure DevOps deliverables and tasks based on data from an Excel file. It generates tasks and deliverables under a specified scenario and assigns tasks to contributors, ensuring efficient management of work items.

#### Features
- Automated Work Item Creation
- Task Assignment
- Scenario-Based Organization

### 2. ADO Scenario Overview

The **ADO Scenario Overview** project generates comprehensive scenario reports for contributors and deliverables, providing summaries and detailed information in a well-organized format.

#### Features
- Scenario Summary
- Sprint Overview
- Detailed Deliverable Sheets
- Deliverable Details

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

3. **Build the Solution**
    ```bash
    dotnet build
    ```

## Configuration

Each project uses a `Config` class for configuration. Update the following parameters in `Config.cs` according to your requirements:

- **ADO Scenario Planner**
  - `PersonalAccessToken`: Token for authenticating API requests.
  - `ScenarioID`: ID of the scenario for task generation.
  - `Organization`: Azure DevOps organization name.
  - `Project`: Azure DevOps project name.
  - `ExcelFilePath`: Path to the Excel file for input data.

- **ADO Scenario Report Generator**
  - `PersonalAccessToken`: Token for authenticating API requests.
  - `ScenarioID`: ID for the specific scenario to generate reports.
  - `SpecialContributors`: Set of contributors with special capacity multipliers.
  - `CapacityMultiplierSpecial`: Multiplier for special contributors.
  - `CapacityMultiplierRegular`: Multiplier for regular contributors.
  - `EndDate`: End date for the reporting period.
  - `CurrentDate`: Current date for generating the report.
  - `Organization`: Azure DevOps organization name.
  - `Project`: Azure DevOps project name.
  - `TasksBaseUri`: Base URI for Azure DevOps task API.
  - `ExcelFilePath`: Path and name template for the generated Excel file.

## Usage

### ADO Scenario Planner

After configuring the tool and ensuring the Excel file is ready, run the following command:

```bash
dotnet run --project <path_to_ADOScenarioPlanner_project>
```

### ADO Scenario Overview
To generate reports, run the following command:

```bash
dotnet run --project <path_to_ADOScenarioReportGenerator_project>
```