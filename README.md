# mailgsheet - Sends data with your Google Sheet

mailgsheet is an application that allows you to easily send row data extracted from a Google Sheet document for every e-mail in a user defined column.

## System Requeriments

- Python 3 (tested with 3.7)

## Quick Start

- [Enable Google Sheets API access](https://developers.google.com/sheets/api/quickstart/python) and download credentials (credentials.json) and put on same folder

- [Fill up mail_credentials_example.json](https://github.com/eduardomarossi/mailgsheet/blob/master/mail_credentials_example.json) and rename to mail_credentials.json in same folder

- Use cases:

   - Sends every contact (column e-mail) the row data from a sheet named Page1 from column A to E. The header of sheet is on line 1 and data starts in line 2. Adds your manager contact on copy.

        ``` bash
        python3 main.py https://docs.google.com/spreadsheets/d/XXXXXXXXXXX/edit Page1 A:E --mail-column "e-mail" --header-lines 1 --rows-start 2 -add-cc manager@mycompany.com 
        ```

## Command Line Usage

| Parameter | Description | Example |
| --------  | ------------ | ------- |
| sheet_url | Google Sheet URL | http://sheets/d/XXXXXXXXXXXX/edit?usp=sharing
| sheet_name | Sheet name | Page1 |
| sheet_range | Range where all data is contained | A:N - A column to N (all rows)
| --header-lines | Specify line number where header is contained | 1-3 or 1 |
| --rows-start | Specify where the table data starts (default: line after header) | 4
| --mail-column | Name of the column where e-mails are stored (one per row) | Mail
| --dry-run | Output resulting e-mails in screen and don't send | |
| --mail-credentials-path | Specify path to mail credentials (for sending mails) | Check mail_credentials_example.json
| --google-credentials-path | Google App Credentials Json file | |
| --debug | Enable debug printing | |
| --verbose | Enable debug printing | |
| --debug-force-to | Forces all mail be sent to specified e-mail | myemail@provider.com
| --add-cc | Adds an e-mail to cc field (copy) | manager@mycompany.com |
| --debug-send-interval-start | Sends only mails in specified interval | 3 (first three e-mails are ignored)
| --debug-send-interval-end | Sends only mails in specified interval | 7 (sends only up to seventh e-mail).



## License

GPL v2.0
