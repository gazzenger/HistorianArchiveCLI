# HistorianArchiveCLI
*For GE Proficy Historian (currnently works with v7.1 and v8.1 Historian)*
**v 2.0.0.0**
**By Gary**

This is CLI for GE Proficy Historian to bulk restore archives. This application requires the Historian SDK installed.

```bash
usage: HistorianArchiveCLI [-s servername] [-u username] [-p password] [--restore] OR [--backup] AND/OR [--remove] [-f filename] [-olderthan days] [-exportpath path] [-o] [-d datastore] [-nc]

        -s      The hostname of the server for connecting to
        -u      The username for authenticating with historian (leave blank to use Windows Authentication)
        -p      The password for authenticating with historian (leave blank to use Windows Authentication)
        --restore       Restoring archives listed in the file
        --backup        Backing up archives, either from a file or based on an age
        --remove        Removing archives, either from a file or based on an age
        -f      The file path to the file containing the list of the archive paths to import for restoring (only IHAs at the moment), one item per line
        -olderthan      Archives older than a specified age (in days from today) to backup and/or remove
        -exportpath     The path to move backed up or removed archives
        -o      Overwrite any existing IHA files in the default archive path
        -d      The specified datastore (leave blank to use the default datastore)
        -nc     Skip the backing up of configuration file (IHC) before restoration
```

### Requirements
- Visual Studio (or equivalent) - this was built using Visual Studio 2019 as a .NET Framework VB Console Application

NB the Historian SDK *should* work on C# as well as VB, and may also work with .NET Core.

- Historian SDK
Install this from the Historian installation disk, under Client Tools

The ISO can be downloaded (with an account) from the GE website, https://digitalsupport.ge.com/en_US/Download/Historian-7-1


### Inputs
#### File List
A file list should contain one IHA file per line
```text
C:\...\asdf.iha
C:\...\asdf2.iha
...
```
#### Olderthan
An older than flag can be used to define archives older than a certain number of days to be selected.

### Outputs
If using backup or remove, provide an export path, to move the archives after the operation.

### Historian SDK documentation

The main documentation can be found at the following links,
- [Predix Website](https://docs.predix.io/en-US/content/historian/apis_and_sdk/historian_sdk/)
- [GE Website](https://www.ge.com/digital/documentation/historian/version71/IMGI4YzMyN2EtM2JhMC00NmQ2LTg2N2MtYjQ2NGE0ZTlhNjhh.html#IMGI4YzMyN2EtM2JhMC00NmQ2LTg2N2MtYjQ2NGE0ZTlhNjhh)


### To be added
- An actual CLI, with the ability to run simple commands within the application (rather than simply running the command with inputs)
- Considering adding a secondary library to utilise the Client Access Assembly - this uses WCF on port 13000
