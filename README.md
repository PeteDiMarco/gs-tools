# Miscellaneous Google Apps Script Tools

This is a collection of functions I wrote while developing a complex spreadsheet in
Google Sheets.  They can be split into 3 categories:
1. *Utility Functions:* These functions compensate for features missing from Google Apps
Script.
2. *Formula Functions:* Functions that can be used in Google Sheets formulas.
3. *Class Constructor Functions:* An "Enum" class.

[More](docs/index.html)

## Disclaimer

Google Apps Script is a subset of JavaScript 1.6.  Some features of JS 1.6 (such as
"let" and "const") are either buggy or absent.  Since I'm not an expert JavaScript 
programmer, I wrote this code to work around those limitations.  There may be more
conventional solutions to these problems, especially my Enum class.  I also tend to
think in Python, which may cause JavaScript pros some distress.

Some important things to be aware of:
1. Functions with names that end with "\_" cannot be called as formulas from Google Sheets.
2. User-defined formulas that are called too frequently or run too long may be rate-limited
by Google.
3. Google Sheets is not a reasonable platform for complex applications.  Use Rails or 
Django instead.

## Versioning

N/A (yet).

## Authors

* **Pete DiMarco** - *Initial work* - [PeteDiMarco](https://github.com/PeteDiMarco)

## License

This project is licensed under the Apache 2.0 License - see the [LICENSE](LICENSE) file
for details.

