# Google Spreadsheets output plugin for Embulk

## Overview

* **Plugin type**: output
* **Load all or nothing**: no
* **Resume supported**: no
* **Cleanup supported**: no

## Configuration

- auth_method (string, default: 'authorized_user'): 'authorized_user' or 'service_account'
- json_keyfile (string, required): credential file path or `content` string
- spreadsheet_url (string, required): your spreadsheet's url
- worksheet_title (string, required): worksheet's title
- mode (string, default: append_direct): writing record method, available mode are `append_direct` and `delete_in_advance`
- header_line (bool, default: false): if true, write header to first record
- start_row (integer, default: 1): specific the start row
- start_column (integer, default: 1): specific the start column
- null_string (string, default: ''): replace null to `null_string`
- default_timezone (string, default: '+00:00'): time zone offset of timestamp columns
- default_timestamp_format (string, default: '%Y-%m-%d %H:%M:%S.%6N %z'): output format of timestamp columns

**json_keyfile**

specific the credential file which the Google Developer Console provides for authorization.
https://console.developers.google.com/apis/credentials
if use oauth, should be included the `refresh_token` field in credential json file.

```
{
  "client_id": "******************************************************",
  "client_secret": "***************************",
  "refresh_token": "***************************"
}
```

## Example

```
out:
  type: google_spreadsheets
  json_keyfile:
    content: |
      {
        "client_id": "******************************************************",
        "client_secret": "***************************",
        "refresh_token": "***************************"
      }
  # json_keyfile: './keyfile.json'
  spreadsheet_url: 'https://docs.google.com/spreadsheets/d/16RSM_xj5ZB4rz0WBlnIbD1KHO46KASnAY04e_oYUSEE/edit'
  worksheet_title: 'シート1'
  start_row: 1
  start_column: 1
  header_line: true
  null_string : '(NULL)'
  default_timezone: '+09:00'
  default_timestamp_format: "%Y-%m-%d %H:%M:%S %z"
```

## Mode

* append_direct
  * Behavior: This modes appends rows to the target worksheet directly. If the target worksheet doesn't exist, it is created automatically
  * Transactional: No. If fails, the target worksheet could have some rows inserted.
  * Resumable: No.
* delete_in_advance
  * Behavior: Delete rows to be appended first, then `append_direct`.
  * Transactional: No. Users may see all rows are once deleted. Also, if fails, the target worksheet could have some rows inserted.
  * Resumable: No.

## Build

```
$ rake
```
