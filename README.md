# HistorianArchiveCLI
##v 1.0.0.0
##By Gary Namestnik

This is CLI for GE Historian to bulk restore archives. This application requires the Historian SDK installed.

```bash
HistorianArchiveCLI [-s] [-u] [-p] -f [-o] [-d]

        -s      The hostname of the server for connecting to
        -u      The username for authenticating with historian (leave blank to use AD group)
        -p      The password for authenticating with historian (leave blank to use AD group)
        -f      The file path to the file containing the list of the archive paths to import for restoring (only IHAs at the moment), one item per line
        -o      Overwrite any existing IHA files in the default archive path
        -d      The specified datastore (leave blank to use the default datastore)
```

##File List
A file list should contain one IHA file per line

