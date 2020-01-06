# HistorianArchiveCLI
*For GE Proficy Historian (currnently works with v7.1 Historian)*
**v 1.0.0.0**
**By Gary Namestnik**

This is CLI for GE Proficy Historian to bulk restore archives. This application requires the Historian SDK installed.

```bash
HistorianArchiveCLI [-s] [-u] [-p] -f [-o] [-d] [-nc]

        -s      The hostname of the server for connecting to
        -u      The username for authenticating with historian (leave blank to use AD group)
        -p      The password for authenticating with historian (leave blank to use AD group)
        -f      The file path to the file containing the list of the archive paths to import for restoring (only IHAs at the moment), one item per line
        -o      Overwrite any existing IHA files in the default archive path
        -d      The specified datastore (leave blank to use the default datastore)
        -nc     Skip the backing up of configuration file (IHC) before restoration
```

### Requirements
- Visual Studio (or equivalent) - this was built using Visual Studio 2019 as a .NET Framework VB Console Application

NB the Historian SDK *should* work on C# as well as VB, and may also work with .NET Core (I haven't tried it)

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
### Historian SDK documentation

The main documentation can be found at the following links,
- [Predix Website](https://docs.predix.io/en-US/content/historian/apis_and_sdk/historian_sdk/)
- [GE Website](https://www.ge.com/digital/documentation/historian/version71/IMGI4YzMyN2EtM2JhMC00NmQ2LTg2N2MtYjQ2NGE0ZTlhNjhh.html#IMGI4YzMyN2EtM2JhMC00NmQ2LTg2N2MtYjQ2NGE0ZTlhNjhh)


### To be added
- An actual CLI, with the ability to run simple commands within the application (rather than simply running the command with inputs)

