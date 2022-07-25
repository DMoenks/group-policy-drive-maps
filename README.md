# group-policy-drive-maps

⚠️**This repository most probably won't be updated anytime soon, as I don't work with group policies that much as of now**

## Use case description
This script creates or modifies a GPO with drive maps configured in a matching Excel workbook.

## Feature overview

## Preparations
The Excel workbook needs to be configured as follows:
- The worksheet holding the configuration needs to be named 'DriveMaps'
- The first row is ignored and may therefore contain headings
- Starting from the second row the columns need to contain the following information:
    1. UNC path pointing to the share to map
    2. Drive letter
    3. Drive label (optional)
    4. Filter (optional, can either be a group name or a distinguished name pointing at an OU)

## Links
- [Group Policy Drives](https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-gppref/f1df2af2-0189-4da4-9d89-d369064015f7)
- [Group Policy Targetting](https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-gppref/ca4571c2-7906-49fd-bae3-f2e58c098345)
