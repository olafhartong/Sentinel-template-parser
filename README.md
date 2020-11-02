# Sentinel Templates to XLSX

[![license](https://img.shields.io/github/license/olafhartong/sysmon-modular.svg?style=flat-square)](https://github.com/olafhartong/sysmon-modular/blob/master/license.md)
![Maintenance](https://img.shields.io/maintenance/yes/2020.svg?style=flat-square)
[![GitHub last commit](https://img.shields.io/github/last-commit/olafhartong/sysmon-modular.svg?style=flat-square)](https://github.com/olafhartong/sysmon-modular/commit/master)
[![Twitter](https://img.shields.io/twitter/follow/olafhartong.svg?style=social&label=Follow)](https://twitter.com/olafhartong)

This small script converts all template rules obtained from the API to XLSX for easier reference.

# Usage

## Getting the data from the Microsoft API Manually
* The JSON can be retrieved here: https://docs.microsoft.com/en-us/rest/api/securityinsights/alertruletemplates/list

* Then run: ``` ./sentinel-template-parse.ps1```

## Using AzSentinel

Use the great AzSentinel by @pkhabazi and @wortell https://github.com/wortell/AZSentinel
* Install and authenticate the module per the authors instructions

* Edit the Workspace variable in azsentinel-template-parse.ps1

* Then run: ``` ./azsentinel-template-parse.ps1```


Both scripts update the xlsx file