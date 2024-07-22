# Label Generation Script

This script processes a list of orders and generates labels using specified PDF templates and barcodes.

## Table of Contents

- [Label Generation Script](#label-generation-script)
  - [Table of Contents](#table-of-contents)
  - [Features](#features)
  - [Requirements](#requirements)
  - [Setup](#setup)
  - [Configuration](#configuration)
  - [Usage](#usage)
  - [Logging](#logging)
  - [Contributing](#contributing)
  - [License](#license)

## Features

- Generates labels based on PDF templates and barcode data.
- Uses multi-threading for faster processing.
- Handles errors gracefully with comprehensive logging.
- Configurable via a JSON file.

## Requirements

- Python 3.6+
- Required Python packages (see `requirements.txt`)

## Setup

1. **Dowload Python**:
    Go to Microsoft Store and dowload 3.11
    https://www.microsoft.com/store/productId/9NRWMJP3717K?ocid=pdpshare

2. **Open the Terminal**
    Open the Folder what you edited in File Browser and right cklick on a Free place.
    Than please chose "open in terminal"

2. **Install the required packages**:
    pip install -r requirements.txt

3. **Create a `config.json` file**:
    python gemerate.py

## Configuration

Ensure that the `config.json` file is in the root directory of the project and contains the correct paths to the necessary files and folders.

## Usage

Run the script using the following command:
```bash
python generate.py --config config.json
