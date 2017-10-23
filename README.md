# Google Spreadsheets output plugin for Embulk

## Overview

* **Plugin type**: output
* **Load all or nothing**: no
* **Resume supported**: no
* **Cleanup supported**: no

## Configuration

- json_keyfile (string, required): credential file path
- spreadsheet_id (string, required): your spreadsheet's id
- worksheet_gid (integer, default: 0): worksheet's gid if you want to specific worksheet
- mode (string, default: append): writing record method, available mode are `append` and `replace`
- is_write_header_line (bool, default: false): if true, write header to first record
- start_cell (string, default: 'A1'): specific the index of first record by the A1
  notation.
- null_string (string, default: ''): replace null to `null_string`
- default_timezone (string, default: 'UTC'): time zone of timestamp columns

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

**mode**

## Example

```
out:
  type: google_spreadsheets
  json_keyfile: './keyfile.json'
  spreadsheet_id: '16RSM_xj5ZB4rz0WBlnIbD1KHO46KASnAY04e_oYUSEE'
  worksheet_gid: 1519744516
  mode: replace
  start_cell: 'B2'
  is_write_header_line: false
  null_string : '(NULL)'
  default_timezone: 'Asia/Tokyo'
```


## Build

```
$ rake
```
