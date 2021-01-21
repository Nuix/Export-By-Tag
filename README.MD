Export by Tag
=============

[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](http://www.apache.org/licenses/LICENSE-2.0) ![This script was last tested in Nuix 9.0](https://img.shields.io/badge/Script%20Tested%20in%20Nuix-9.0-green.svg)

View the GitHub project [here](https://github.com/Nuix/Export-By-Tag) or download the latest release [here](https://github.com/Nuix/Export-By-Tag/releases).

# Overview

This script assists by performing a series of exports based on a selection of tags. For each selected tag, a PST or NSF mail store is exported containing item's with the given tag.

# Getting Started

## Setup

Begin by downloading the latest release of this code.  Extract the contents of the archive into your Nuix scripts directory.  In Windows the script directory is likely going to be either of the following:

- `%appdata%\Nuix\Scripts` - User level script directory
- `%programdata%\Nuix\Scripts` - System level script directory

# Settings

## Tags Tab

| Setting | Description |
|---------|-------------|
| **Include Nested Tags** | When this option is checked, for each provided tag, each nested sub tag will be included as well. Note that each specified tag results in a separate export. If you select a parent tag and a sub tag, the parent tag will yield an export of the tag and sub tag, the selected sub tag will yield a second export of just the sub tag. |
| **Use Tags from CSV** | When checked the tags to be exported will be based on a user provided CSV. |
| **Input CSV** | The CSV to load tags from when **Use tags from CSV** is checked. |
| **Use Selected Tags** | When checked the tags to be exported will be based on the user's selection in the table listing of tags present in the case. |

## Options Tab

| Setting | Description |
|---------|-------------|
| **Email Format** | Allows you to choose whether the export will be to a **PST** or **NSF**. |
| **Output Directory** | Used to specify the export root directory. |
| **Output CSV** | Used to specify where the script should generate a report about the exports. |
| **Create Default Production Set** | By default Nuix creates a production set for exports. When this option is unchecked the script will disable this default behavior preventing Nuix from creating production sets from the exports. |
| **Scope Query** | Allows you to provide a search query which will be joined with each tag query, allowing you to further scope down the items exported. |

## Worker Settings Tab

Allows you to specify worker settings used by the export.

**Note:** If a value for the argument nuix.processing.sharedTempDirectory was provided to Nuix upon startup, this value will override any value selected by the user in this tab!

# Input CSV Format

| Column | Description |
|--------|-------------|
| **Tag** | The tag to export. |
| **Directory** *(optional)* | If this column is provided it specifies the sub directory this tag will be exported to within the provided export directory. If this column is not provided the tag name will be used (with illegal filesystem characters replaced with underscores). |

# Cloning this Repository

This script relies on code from [Nx](https://github.com/Nuix/Nx) to present a settings dialog and progress dialog.  This JAR file is not included in the repository (although it is included in release downloads).  If you clone this repository, you will also want to obtain a copy of Nx.jar by either:
1. Building it from [the source](https://github.com/Nuix/Nx)
2. Downloading an already built JAR file from the [Nx releases](https://github.com/Nuix/Nx/releases)

Once you have a copy of Nx.jar, make sure to include it in the same directory as the script.

# License

```
Copyright 2021 Nuix

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
```
