# ADO Scenario Planner

## Overview

The **ADO Scenario Planner** tool automates the creation of Azure DevOps deliverables and tasks based on data from an Excel file. It generates tasks and deliverables under a specified scenario and assigns tasks to contributors, ensuring efficient management of work items within Azure DevOps.

## Features

- **Automated Work Item Creation**: Creates deliverables and tasks in Azure DevOps based on the provided Excel data.
- **Task Assignment**: Automatically assigns tasks to contributors as specified in the Excel file.
- **Scenario-Based Organization**: Organizes tasks under specific scenarios and iterations.

## Prerequisites

- .NET 6.0 or later
- [ClosedXML](https://github.com/ClosedXML/ClosedXML) library for Excel file handling

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
    ```

## Configuration

The tool uses the `Config` class for configuration. Update the following parameters in `Config.cs`:

- `PersonalAccessToken`: Token for authenticating API requests (dummy value used in examples).
- `ScenarioID`: ID of the scenario for which tasks and deliverables are to be generated.
- `Organization`: Azure DevOps organization name.
- `Project`: Azure DevOps project name.
- `ExcelFilePath`: Path to the Excel file containing the task and deliverable data (`Deliverables_Tasks_input.xlsx`).

**Example Config Settings:**

```csharp
public static class Config
{
    public static readonly string PersonalAccessToken = "dummy_personal_access_token";
    public static readonly int ScenarioID = 12345678;
    public static readonly string Organization = "exampleorg";
    public static readonly string Project = "exampleproject";
    public static readonly string ApiVersion = "7.0";
    public static readonly string ExcelFilePath = "C:\Deliverables_Tasks_input.xlsx";
}
```

## Usage

### Ensure the Excel File is Ready

The Excel file should contain deliverables and tasks with columns for **Deliverable**, **Task**, **Owner**, **Iteration**, **Original Estimate**, and **Remaining Days**.

Example Excel file structure (`Deliverables_Tasks_input.xlsx`):

| Deliverable         | Task        | Owner            | Iteration      | Original Estimate | Remaining Days |
|---------------------|-------------|------------------|----------------|-------------------|----------------|
| Dummy Deliverable 1  |             | test@test.com    | 2401\2401-2    | 4                 |                |
|                     | Dummy Task1 | test@test.com    | 2401\2401-1    | 1                 | 1              |
|                     | Dummy Task2 | test@test.com    | 2401\2401-2    | 3                 | 3              |
| Dummy Deliverable 2  |             | test2@test.com   | 2401\2401-2    | 4                 |                |
|                     | Dummy Task3 | test2@test.com   | 2401\2401-1    | 1                 | 1              |
|                     | Dummy Task4 | test2@test.com   | 2401\2401-2    | 3                 | 3              |

### Run the Tool

After configuring the tool and ensuring the Excel file is in place, run the following command:

```bash
dotnet run
```
The tool will:

- **Read** the tasks and deliverables from the Excel file.
- **Create** work items in Azure DevOps.
- **Assign** tasks to contributors as specified in the Excel file.
