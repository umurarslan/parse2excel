# Parse2Excel

Simple CLI tool parsing text files with TextFSM and export to SQLite and Excel with own configuration file, also within configuration file custom SQLite commands can be executed to create custom tables.

---

## Requirements

[Python >= 3.9](https://www.python.org/downloads/)

> For Windows, select the **Add Python 3.x to PATH** checkbox during installation.

---

## Installation

**Option 1:**

From Python Package Index (PyPI) Repo:

```
pip install parse2excel
```

**Option 2:**

Download project ZIP file and run below command:

```
pip install parse2excel-X.zip
```

### Installation Check

After installation **parse2excel** command added to **System Path** and can be executed from any path easily as below:

```
> parse2excel -h
usage: parse2excel [-h] [configfile]

positional arguments:
  configfile   config yaml file path [e.g. srlinux_config_1.yaml] (OPTIONAL default: file=config.yaml, folder=P2E_CONFIGS)

options:
  -h, --help  show this help message and exit
```

---

## Usage

### Simple Usage

Run **parse2excel** command with **Config File Path OR without argument "config.yaml" file used (single config file) OR "P2E_CONFIGS" Folder used (multiple config files in folder)** from any path after that check excel/SQLite output files in working directory.

```
parse2excel <Config_File_Path>

```

---

### Config file

There are 4 different options in YAML config file:

- **textfsm:** Parse text files in folder with TextFSM template and export excel/SQLite.
- **sqljoin:** Run "SELECT" SQLite command and export excel/SQLite. (Python function supported)
- **sqlfunction:** Create SQLite Python functions for all **sqljoin** parts.
- **excel:** Import Excel file and convert to SQLite.

---

Example config.yaml file:

```yaml
- type: textfsm
  db_name: my_p2e_excel
  table_name: my_interface_sheet
  # excel_export: none
  folders:
    - config_FOLDER
  template: |
    Value Required Interface (\S+)
    Value Interface_Description (\S+)
    Value Interface_Ip (\S+)
    Value Interface_Mask (\S+)

    Start
      ^interface ${Interface} -> Begin

    Begin
      ^ description ${Interface_Description}
      ^ ipv4 address ${Interface_Ip} ${Interface_Mask}
      ^! -> Record Start

- type: sqljoin
  db_name: my_p2e_excel
  new_table: SecGW_CERT
  sqlcommand: select removetxt(Filename), Cert_Name, Cert_Start, Cert_End from certificate
  functions:
    - |
      def removetxt(d):
        return d.replace('.txt','')

- type: sqlfunction
  functions:
    - |
      def removetxt(d):
          # add function to all sqljoin parts
          return d.replace('.txt','')

- type: sqljoin
  db_name: my_p2e_excel
  # only run any sqlite command for debug, delete table etc.
  sqlcommand_run: select Service_ID from vprn_w_x

- type: sqljoin
  db_name: cer
  new_table: cer_int_with_policy
  first_table: cer_interface
  second_table: cer_vrf_policy
  match: Vrf
  # match: cer_interface.vrf1 = cer_vrf_policy.vrf2

- type: excel
  db_name: from_excel
  excel_file: excel_file.xlsx
  # default ALL OR OPTIONAL excel_sheets
  excel_sheets:
    - Sheet1
```

---

**_TBA_**
