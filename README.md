# PagerDuty-Report-Powershell

PagerDuty report that is used to fetch current acknowledged, triggered and high severity alerts and display them into a readable HTML table. This report can be either be emailed or used as a standalone HTML file.

## Getting Started

### Installation

You will need to have PowerShell installed on your machine.

Windows - Installed by default although there might be some policy restriction settings that prevent you from running scripts. You may need to run the following in an Admin PowerShell console: <br>
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
```

MacOS - You will to install the [Homebrew]() package manger by running the following in your Terminal: <br>
`/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/master/install.sh)"`

Then install PowerShell by running the following in your Terminal: <br>
`brew cask install powershell`

### Usage

You will need to include in your own API as well as specific Team ID:
```powershell
# API Token
$apiKey = "API KEY"

# URI for Triggered PD alerts for a specifc Team
$URI_triggered = "https://api.pagerduty.com/incidents?statuses[]=triggered&team_ids[]=TEAM_ID"

# URI for Acknowledged PD alerts for a specifc Team
$URI_ack = "https://api.pagerduty.com/incidents?statuses[]=acknowledged&team_ids[]=TEAM_ID"
```

### Contributing

1. Fork the master branch
2. Create your feature branch: `git checkout -b my-new-feature`
3. Commit your changes: `git commit -am 'Add some feature'`
4. Push to the branch: `git push origin my-new-feature`
5. Submit a pull request!
---

### Author/ Contact
**Ayush Lal** <br>
hello@ayushlal.com.au <br>
[Portfolio website](http://www.ayushlal.com.au) <br>
[GitHub](https://github.com/ayush-lal)

*// If you have any queries please feel free to get in touch.*

