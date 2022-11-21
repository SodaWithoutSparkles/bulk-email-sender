# bulk-email-sender
A script for sending multiple emails at once! Or you could say its a email merge...

## Features
- Easy config
- Compatible with Excel
- Heavily commented
- Support multiple links and variables

## Dependencies
- Python3 ([install](https://www.python.org/downloads/))
- `pandas`
- `openpyxl`

## Installation
### Method 1 (git clone)
1. `git clone https://github.com/SodaWithoutSparkles/bulk-email-sender.git`
2. Navigate to `bulk-email-sender`

### Method 2 (zip)
1. On this github page, click the green `<Code>` icon, then download as zip
2. Extract the zip file

## Usage
1. Use your choice of editor to edit the config section of the code
2. Install python from [python.org](https://www.python.org/downloads/)
3. Tick "add python to PATH" when installation [as shown](https://docs.blender.org/manual/en/latest/_images/about_contribute_install_windows_installer.png) (credit: blender)
4. From a shell/cmd, execute this script by `python3 <path-to-the-script>`

## Excel format
| name  | to                | links                                  |
|-------|-------------------|----------------------------------------|
| alice | alice@example.com | https://example.com https://google.com |
| bob   | bob@example.net   | https://example.net                    |
| chris | chris@example.org | https://example.org                    |

## License
[MIT License](https://github.com/SodaWithoutSparkles/bulk-email-sender/blob/main/LICENSE)
