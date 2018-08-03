# Inboxscanner

Scans a configured outlook inbox folder with yara rules, a list of IOCs and external api scanners (VirusTotal,etc...).
Results are attached to the email. 

## Dependencies

pywin32,requests,unidecode and yara-python are required. 

The win32 directory has the latest stand-alone exe file which does not need any dependencies satisfied.

# Usage

Configure config.json and place it in the same directory as the python script or .exe file.

