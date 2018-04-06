## Python Excel Macros

Tools that enable a Python based scripting environment using [xlwings](https://www.xlwings.org/).

A `.xlsm` compiler script is available to automatically create a `PERSONAL.xlsm` file that calls references to python scripts defined in `config.json`.

This project serves two major purposes:
* Use Python instead of VBA for Excel macros to ease development and add source control to scripts
* Eliminate the headache of manually editing the `PERSONAL.xlsm` file when adding new xlwings scripts that use `RunPython`