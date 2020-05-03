# NSGrowth

Usage
Requires 2 sets of nation and region data dumps. Uses the NS Datadump archive format (yyyy-mm-dd-<dumptype>.xml.gz)
If downloading from the archive, ensure you replace -xml.gz with .xml.gz, and place all data dumps in a folder named `dumps`
`dotnet ./DatadumpTool.dll <date1> <date1>`
Example:
`dotnet ./DatadumpTool.dll 2020-04-17 2020-04-18`

The resulting spreadsheet will be placed in `sheets` directory.