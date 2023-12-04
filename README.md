# GNC SQL_Data/Dashboard

A versatile R script for automating the download, processing, and organization of email attachments from a Gmail account.

## Table of Contents
- [Introduction](#introduction)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [Folder Structure](#folder-structure)
- [Contributing](#contributing)
- [License](#license)

## Introduction

## Automated Reporting Workflow

### ThursdayReport R Script

The `ThursdayReport` R script is scheduled to run every Thursday via a .bat file. This script identifies the latest email within the specified email account, meeting specific subject and sender criteria. It then downloads all attached files, performs cleaning and transformation processes, and finally saves the processed data to a designated folder with a custom name. If there are existing files with the same name, the script overwrites them.

### DailyTransaction R Script

The `DailyTransaction` R script runs daily, triggered by a .bat file. It extracts data from attached files and saves it to a specific folder, overwriting data from three days prior.

### SQL Server Integration Services (.dtsx)

The `.dtsx` script plays a crucial role in updating and uploading data to the SQL server of the brand thorugh a staging tables proccess. The data loaded into the server is automatically processed thanks to the previously mentioned R scripts, ensuring seamless integration with the corporate data source and maintaining data integrity.

These automated scripts collectively create an efficient workflow, enabling regular data processing, transformation, and update tasks with minimal manual intervention.


## Features

- **Flexible Querying**: Use Gmail search queries to filter emails based on specific criteria such as sender, subject, or date.
- **Organized Storage**: Automatically create folders for different types of attachments and organize them according to your specified criteria.
- **File Processing**: Includes functionality to split and process various types of attachments, making it easy to work with different file formats.

## Prerequisites

Before using the script, ensure you have the following installed:

- R (https://www.r-project.org/)
- Required R packages: `gmailr`, `filesstrings`, `readxl`, `readr`, `writexl`

## Installation

1. Install R from [https://www.r-project.org/](https://www.r-project.org/).
2. Install required R packages:

    ```R
    install.packages(c("gmailr", "filesstrings", "readxl", "readr", "writexl"))
    ```

## Configuration

1. Open the R script in your preferred code editor.
2. Set your Gmail account credentials:

    ```R
    gmail_address <- "your_email@gmail.com"
    gmail_password <- "your_password"
    ```

3. Configure the Gmail authentication:

    ```R
    gm_auth_configure(path = "path/to/gargle/token/folder")
    gm_auth(path = "path/to/gargle/token/folder")
    ```

4. Specify your Gmail query, destination folders, and file names:

    ```R
    query <- "from:sender@example.com subject:your_subject"
    destination_folders <- c("folder1", "folder2", "folder3")
    destination_file_names <- c("file1.xlsx", "file2.xlsx", "file3.xlsx")
    ```

## Usage

1. Run the script using your preferred R environment.
2. The script will authenticate with Gmail, download attachments based on your criteria, and organize them into specified folders.

## Folder Structure

- **attachments**: Folder to store downloaded attachments.
- **folder1, folder2, folder3**: Destination folders for different types of reports.

## Contributing

Contributions are welcome! If you have any suggestions, feature requests, or find a bug, please open an issue or create a pull request.

## License

This project is licensed under the [MIT License](LICENSE).
