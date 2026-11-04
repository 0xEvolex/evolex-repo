# evolex-repo

Release repository for Windows tools by **Evolex** for Silkroad Online and Sailor Online.
It holds the published builds and the data files the tools download at runtime. Source code lives
in separate repositories.

## Applications

| Application | Version | Description | Download |
| --- | --- | --- | --- |
| Sailor Events | 2.5.0.0 | Event schedule and reminders for Sailor Online | [Sailor Events.exe](https://raw.githubusercontent.com/0xEvolex/evolex-repo/main/sailor-events/Sailor%20Events.exe) |
| Sailor Fixer | 2.1.0.0 | Diagnostic and repair toolkit for the Sailor Online client and common Windows issues | [Sailor Fixer.exe](https://raw.githubusercontent.com/0xEvolex/evolex-repo/main/sailor-fixer/Sailor%20Fixer.exe) |
| SRO Hider | 1.3.0.0 | Hides and restores Silkroad client windows | [SRO Hider.exe](https://raw.githubusercontent.com/0xEvolex/evolex-repo/main/sro-hider/SRO%20Hider.exe) |
| SRO Sleep | 1.1.2.0 | Raises an alarm on disconnect, character death or client crash | [SRO Sleep.exe](https://raw.githubusercontent.com/0xEvolex/evolex-repo/main/sro-sleep/SRO%20Sleep.exe) |
| SRO Alchemist | 0.0.0.8 | Alchemy automation for Silkroad Online | [SRO Alchemist.exe](https://raw.githubusercontent.com/0xEvolex/evolex-repo/main/sro-alchemist/SRO%20Alchemist.exe) |
| SRO Opus | 0.0.4.2 | Automation for Silkroad Online | [SRO Opus.exe](https://raw.githubusercontent.com/0xEvolex/evolex-repo/main/sro-opus/SRO%20Opus.exe) |

All applications run on Windows 10 and 11 (64-bit).

## Sailor Companion modules

`sailor-companion/` is the module catalog for **Sailor Companion**, a hub app that installs only the
tools a player uses. The app reads `catalog.json`, downloads the selected package and checks it
against the SHA-256 listed in the catalog before installing it.

| Module | Version | Package |
| --- | --- | --- |
| Sailor Events | 3.1.0 | `events-3.1.0.zip` |
| SRO Hider | 2.0.1 | `hider-2.0.1.zip` |
| SRO Sleep | 2.1.0 | `sleep-2.1.0.zip` |
| Sailor Fixer | 2.1.0 | `fixer-2.1.0.zip` (the Sailor Fixer release above, unchanged) |

Packages are installed through Sailor Companion; they are not meant to be unpacked by hand.

## Repository layout

| Path | Contents |
| --- | --- |
| `manifest.json` | Name, version and download link of each application |
| `sailor-events/` | Application and the event data it downloads (`events.json`, `sailor_ad_events.json`, `tooltips.yaml`) |
| `sailor-fixer/` | Application, its configuration files and `latest_game_files/` used to restore missing client files |
| `sailor-companion/` | Module catalog and packages for Sailor Companion |
| `sro-hider/`, `sro-sleep/`, `sro-alchemist/`, `sro-opus/` | Application builds |

## Security

Download the tools only from this repository or from links that point to it. Windows SmartScreen or
antivirus software may warn about new, unsigned releases. If a file looks suspicious, do not run it
and report it through the support channels below.

## Support

Report problems through the Sailor Online community channels. Please include the application name,
the version and, for Sailor Fixer, the support file it creates.

## License

Licensed under [Creative Commons BY-NC-ND 4.0](LICENSE). You may share the files unchanged and
with attribution, but not for commercial use and not in modified form.

Silkroad Online and Sailor Online are the property of their respective owners. These tools are
independent and not affiliated with or endorsed by them.
