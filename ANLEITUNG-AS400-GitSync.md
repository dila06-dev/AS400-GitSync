# IBM-i-Quellen automatisch exportieren und nach Git übertragen

Enthalten sind `Sync-AS400SourcesToGit.ps1` und `AS400-GitSync.psd1`. Das vorhandene Skript `SRC_MEMBER_LAST_CHANGE_1DAY.ps1` wird weiterverwendet. Es ist nicht im Paket enthalten.

## Ablauf

1. Git-Verzeichnis, Branch und Index prüfen; parallele Läufe sperren.
2. Vorherige CSV ins Log-Verzeichnis verschieben und die vorhandene API-Abfrage starten.
3. Neue CSV mit Semikolon als Trennzeichen einlesen; doppelte Member entfernen.
4. CSV-Pfade und `CPYTOSTMF`-Parameter prüfen.
5. Bei `-Execute`: Verbindung über IBM i Access ODBC aufbauen. `CALL QSYS2.QCMDEXC(?)` führt die CL-Befehle synchron aus.
6. Die Zuordnung des IFS-Pfades zum Windows-Pfad mit einer einmaligen temporären Exportdatei prüfen; diese danach wieder entfernen.
7. Alle Member mit UTF-8 / CCSID 1208 und LF exportieren. Fehlende Unterordner werden über den Windows-Zugriff angelegt.
8. Bei einem separaten lokalen Git-Clone die exportierten Dateien kopieren und SHA-256 vergleichen.
9. Nur die CSV-Zieldateien mit `git add` vormerken. Bei tatsächlichen Änderungen einen Commit anlegen.
10. `git push origin HEAD:refs/heads/main` ausführen. Ohne neue Änderungen wird ebenfalls gepusht, damit ein zuvor fehlgeschlagener Push nachgeholt wird.

Der automatische Committext lautet beispielsweise:

```text
IBM i Quellen: 1 Datei(en), 2026-10-01 09:30:00
```

Ein Git-Push überträgt auch andere bereits vorhandene, noch nicht gepushte Commits desselben Branches. Er ist nicht auf den gerade erzeugten Commit beschränkt.

## 1. Dateien ablegen und Konfiguration anpassen

Das Paket nach `D:\AS400\AS400-GitSync` entpacken. `AS400-GitSync.psd1` mit einem Texteditor öffnen.

Das vorhandene `SRC_MEMBER_LAST_CHANGE_1DAY.ps1` ebenfalls in diesen Ordner kopieren. Seine CSV-Ausgabe muss unter `D:\AS400\AS400-GitSync\SRC_MEMBER_LAST_CHANGE_1DAY.csv` liegen. Falls im vorhandenen Abfrageskript ein absoluter Ausgabepfad steht, diesen dort ebenfalls anpassen. Das bestehende Abfrageskript ist nicht im Paket enthalten.

Eine bereits angelegte `ibmi-credential.xml` auf demselben Rechner für denselben Windows-Benutzer ebenfalls in diesen Ordner kopieren; andernfalls wie unten beschrieben neu erstellen.

| Einstellung | Vorbelegung / Bedeutung |
|---|---|
| `QueryScript` | `D:\AS400\AS400-GitSync\SRC_MEMBER_LAST_CHANGE_1DAY.ps1` |
| `CsvPath` | `D:\AS400\AS400-GitSync\SRC_MEMBER_LAST_CHANGE_1DAY.csv` — muss dem tatsächlichen Ausgabeort entsprechen |
| `CsvEncoding` | `utf-8`; bei einer ANSI-Datei z. B. `windows-1252` |
| `IbmSystem` | `s105dd7a.dometic.internal` — prüfen |
| `OdbcDriver` | `IBM i Access ODBC Driver` — mit installiertem Treiber abgleichen |
| `OdbcDsn` | Leer: Driver/System verwenden. Alternativ vorhandenen System-DSN eintragen |
| `CredentialPath` | `D:\AS400\AS400-GitSync\ibmi-credential.xml` |
| `IfsRoot` | `/home/langlitz/AS400` |
| `IfsWindowsRoot` | `F:\AS400` — Windows-Zugriff auf exakt diesen IFS-Ordner |
| `RepoPath` | `F:\AS400` — vorhandenes Git-Arbeitsverzeichnis |
| `Remote` / `Branch` | `origin` / `main` |
| `GitExe` | `git.exe`; falls nicht im PATH, vollständigen Programmpfad eintragen |
| `LogDirectory` | `D:\AS400\AS400-GitSync\Logs` |
| `QueryTimeoutSec` | 900 Sekunden für den gesamten Abfrageprozess |
| `CommandTimeoutSec` | 600 Sekunden pro ODBC-Befehl |
| `GitTimeoutSec` | 300 Sekunden pro Git-Aufruf |

**Die wichtigste Zuordnung:**

```text
IFS:     /home/langlitz/AS400/TVRELWAECO/QRPGLESRC/AKTIVBANKR.sqlrpgle
Windows: F:\AS400\TVRELWAECO\QRPGLESRC\AKTIVBANKR.sqlrpgle
```

Wenn `F:\AS400` bereits direkt auf den IFS-Ordner zeigt, bleiben `IfsWindowsRoot` und `RepoPath` identisch. Das Skript prüft diese Zuordnung bei einem echten Export automatisch. Es legt dazu eine Datei `.as400-sync-probe-<Zufallswert>.tmp` an und entfernt sie wieder. Bei einer unterbrochenen Verbindung kann eine solche Prüfdatei übrig bleiben; sie wird nicht zum Commit hinzugefügt.

Wenn `F:\AS400` nur ein lokaler Clone ist, muss `IfsWindowsRoot` stattdessen auf den tatsächlichen NetServer-Zugriff zeigen. Beispiel mit einem separat vorhandenen Clone:

```powershell
IfsWindowsRoot = '\\s105dd7a.dometic.internal\DEINE_FREIGABE\AS400'
RepoPath       = 'D:\Git\AS400'
```

`DEINE_FREIGABE` ist ein Platzhalter. Freigabename und Unterordner hängen von eurer IBM-i-NetServer-Konfiguration ab. Es wird keine Freigabe eingerichtet und keine Laufwerksverbindung erstellt. Ein konfigurierter ODBC-DSN muss zu derselben IBM i zeigen wie die IFS-Freigabe.

## 2. ODBC und Zugangsdaten vorbereiten

Unter Windows PowerShell nach installierten IBM-Treibern sehen:

```powershell
Get-OdbcDriver | Where-Object Name -Like '*IBM*'
```

Die Architektur muss passen: 64-Bit-PowerShell benötigt einen 64-Bit-ODBC-Treiber. Unterstützt wird Windows PowerShell 5.1; PowerShell 7 unter Windows kann ebenfalls verwendet werden.

Zugangsdaten einmalig unter dem Windows-Benutzer speichern, der das Skript später ausführt:

```powershell
Get-Credential -Message 'IBM-i-Benutzer für den Quellenexport' |
    Export-Clixml -LiteralPath 'D:\AS400\AS400-GitSync\ibmi-credential.xml'
```

Hier IBM-i-Benutzer und Passwort eingeben. Unter Windows schützt `Export-Clixml` das Passwort über DPAPI. Es ist nur unter demselben Windows-Benutzer auf demselben Rechner wieder entschlüsselbar. Die Datei gehört außerhalb des Git-Repositories. Das Passwort wird nicht als Prozessargument weitergegeben.

Der IBM-i-Benutzer benötigt Zugriff auf die Quellmember, `QSYS2.QCMDEXC`, `QSYS/CPYTOSTMF` und Schreibrechte für die IFS-Ziele; für die Prüfdatei außerdem `QSYS/RMVLNK`. Der Windows-Zugriff auf die IFS-Freigabe verwendet die dort bereits eingerichteten SMB-Zugangsdaten. Diese können von den ODBC-Zugangsdaten abweichen.

Optional lassen sich vorhandene ODBC-Einstellungen in `OdbcOptions` ergänzen. Beispielsweise `@{ SSL = '1' }`, wenn TLS auf Client und IBM i eingerichtet ist. Das Skript verändert keine Zertifikats- oder Serverkonfiguration.

Git muss unter demselben Windows-Benutzer bereits eingerichtet sein:

```powershell
git -C 'F:\AS400' status
git -C 'F:\AS400' remote -v
git -C 'F:\AS400' branch --show-current
git -C 'F:\AS400' config user.name
git -C 'F:\AS400' config user.email
```

Das Repository braucht mindestens einen Commit. Für automatische Pushes müssen GitHub-Anmeldedaten bereits über den vorhandenen Credential Manager oder SSH eingerichtet sein. Interaktive Git-Anmeldedialoge werden im Skript unterbunden.

## 3. Testlauf

Zuerst mit der bereits vorhandenen CSV prüfen:

```powershell
& 'D:\AS400\AS400-GitSync\Sync-AS400SourcesToGit.ps1' -SkipQuery
```

Dieser Lauf prüft CSV und Git und zeigt die geplanten CL-Befehle. Er führt keine IBM-i-Befehle, Quelldateiexporte, Commits oder Pushes aus. Er benötigt noch keine IBM-i-Credential-Datei. Logs und eine lokale Sperrdatei werden angelegt; die ODBC-Verbindung und die IFS-Zuordnung werden erst bei `-Execute` geprüft.

Testlauf einschließlich einer neuen API-Abfrage:

```powershell
& 'D:\AS400\AS400-GitSync\Sync-AS400SourcesToGit.ps1'
```

Hier wird die bisherige CSV archiviert und die API-Abfrage tatsächlich gestartet. Der restliche Ablauf bleibt im Testmodus.

## 4. Vollständigen Ablauf starten

```powershell
& 'D:\AS400\AS400-GitSync\Sync-AS400SourcesToGit.ps1' -Execute
```

Mit eigener Commitbeschreibung:

```powershell
& 'D:\AS400\AS400-GitSync\Sync-AS400SourcesToGit.ps1' `
    -Execute `
    -CommitMessage 'AKTIVBANKR: Verarbeitung angepasst'
```

Export und Commit ausführen, Push auslassen:

```powershell
& 'D:\AS400\AS400-GitSync\Sync-AS400SourcesToGit.ps1' -Execute -NoPush
```

Der Rückgabecode ist `0` bei Erfolg und `1` bei einem Fehler. `$LASTEXITCODE` zeigt den Rückgabecode nach dem Aufruf an.

## 5. Vertrag mit dem vorhandenen Abfrageskript

Die CSV muss bei jedem erfolgreichen Aufruf neu geschrieben werden. Ohne Treffer genügt diese Kopfzeile:

```csv
"Zeitstempel";"SRC_LIB";"SRC_FILE";"SRC_MEMBER";"SRC_IFS_PATH";"ZIELPFAD";"COMMAND"
```

Fehlt die neue CSV oder ist sie vollständig leer, bricht der Ablauf ab. Alte Daten werden nicht stillschweigend weiterverwendet. Die vorherige CSV liegt als `previous-*.csv` im Log-Verzeichnis. CSV und Logs werden nicht automatisch gelöscht.

Das Skript verarbeitet QSYS-Quellmember entsprechend deinem Beispiel. `SRC_IFS_PATH` muss leer sein. Der Zielpfad muss unter `IfsRoot` liegen und dem Muster `Bibliothek/Quelldatei/Member.Endung` entsprechen. Zulässig ist `CPYTOSTMF` mit genau diesen Parametern, in beliebiger Reihenfolge:

```text
FROMMBR('/QSYS.LIB/<LIB>.LIB/<FILE>.FILE/<MEMBER>.MBR')
TOSTMF('<ZIELPFAD>')
CVTDTA(*AUTO)
STMFCCSID(1208)
ENDLINFMT(*LF)
STMFOPT(*REPLACE)
```

Ein API-Fehler muss als PowerShell-Fehler oder als Rückgabecode ungleich 0 erkennbar sein. Der neue Wrapper wertet außerdem aufgelaufene PowerShell-Fehler aus. Falls das vorhandene Skript Fehler abfängt, sollte es im Fehlerfall mit `throw` oder `exit 1` enden.

Ein mögliches Muster in dessen vorhandenem `catch`-Block:

```powershell
catch {
    # Vorhandene Diagnoseausgaben können davor stehen.
    throw
}
```

`QueryTimeoutSec = 900` verlängert nur die Wartezeit des Wrappers. Der HTTP-Timeout im vorhandenen Skript muss weiterhin passend gesetzt sein, beispielsweise `-TimeoutSec 660` bei Windows PowerShell 5.1. PHP-, Webserver- und Datenbankgrenzen bleiben davon unabhängig.

`1DAY` begrenzt die fachliche Auswahl weiterhin auf das von der API verwendete Zeitfenster. Nach längeren Ausfällen muss die API-Abfrage für den fehlenden Zeitraum nachgeholt werden; der Wrapper erweitert dieses Fenster nicht.

## 6. Fehler und Wiederanlauf

| Situation | Verhalten / nächster Schritt |
|---|---|
| Keine CSV-Zeilen | Kein Export, kein neuer Commit; mit `-Execute` wird ein ausstehender Push versucht |
| Dateien inhaltlich unverändert | Kein leerer Commit; Push wird trotzdem versucht |
| API-Fehler / Timeout | Abbruch vor Export und Git-Schreibaktionen |
| Ungültiger CL-Befehl / falscher Zielpfad | Abbruch vor ODBC und Export |
| Ein Export schlägt fehl | Abbruch vor `git add`, Commit und Push. Zuvor exportierte IFS-Dateien bleiben bestehen |
| Zieldatei bereits verändert/unversioniert | Abbruch vor Überschreiben; nach Prüfung ggf. `-AllowModifiedTargets` verwenden |
| Bereits gestagte Änderungen | Abbruch. Eigene Änderungen zuerst abschließen oder gezielt aus dem Index nehmen |
| Commit schlägt fehl | Kein Push. Gestagte Exportdateien bleiben zur Prüfung erhalten |
| Push schlägt fehl | Lokaler Commit bleibt erhalten; nächster erfolgreicher Lauf versucht den Push erneut |
| Remote-Branch hat zusätzliche Commits | Normaler Push kann abgelehnt werden; Branch manuell synchronisieren. Kein automatischer Pull, Rebase oder Force-Push |
| Ein anderer Sync läuft | Abbruch durch exklusive Sperrdatei |
| CSV `.sqlrpgle`, Git `.SQLRPGLE` für dieselbe Datei | Abbruch wegen unterschiedlicher Schreibweise. Export-/CSV-Pfad an den bereits versionierten Pfad anpassen |
| `CPDA08A` beim Export | Vorhandene IFS-Datei hat ein anderes CCSID-Attribut. Betroffene Datei separat prüfen und korrigieren; keine automatische Löschung oder Attributänderung |

Nach einem teilweise erfolgreichen Export können einige Ziele bereits verändert sein. Nach Prüfung lässt sich dieselbe CSV erneut verwenden:

```powershell
& 'D:\AS400\AS400-GitSync\Sync-AS400SourcesToGit.ps1' `
    -Execute -SkipQuery -AllowModifiedTargets
```

`-AllowModifiedTargets` erlaubt ausdrücklich das Überschreiben der CSV-Zieldateien mit dem IBM-i-Export. Andere Dateien werden nicht zum Commit hinzugefügt. Bereits gestagte Änderungen bleiben ein Abbruchgrund. Das Skript führt keinen automatischen Rollback von Exportdateien aus.

Für den Ablauf empfiehlt sich ein eigenes Git-Arbeitsverzeichnis. Während eines Laufs dort keine manuellen Git-Aktionen ausführen. Die Sperrdatei verhindert parallele Skriptläufe; sie sperrt nicht andere Programme.

## 7. LF/CRLF-Warnung

Die Meldung `LF will be replaced by CRLF` ist eine Git-Warnung zur Zeilenendenkonfiguration. Die Exporte verwenden bereits `ENDLINFMT(*LF)`.

Um LF für IBM-i-Quellen dauerhaft im Repository festzulegen, können diese Regeln in die vorhandene `.gitattributes` ergänzt werden:

```gitattributes
*.[sS][qQ][lL][rR][pP][gG][lL][eE] text eol=lf
*.[rR][pP][gG][lL][eE] text eol=lf
*.[cC][lL][lL][eE] text eol=lf
*.[rR][pP][gG] text eol=lf
*.[cC][lL][pP] text eol=lf
*.[pP][fF] text eol=lf
*.[lL][fF] text eol=lf
*.[dD][sS][pP][fF] text eol=lf
*.[pP][rR][tT][fF] text eol=lf
```

Diese Änderung separat prüfen und committen. Der Wrapper verändert weder `.gitattributes` noch `core.autocrlf` und führt keine automatische Gesamtnormalisierung aus. Eine bereits abweichend versionierte Datei kann beim ersten normalisierten Commit viele geänderte Zeilen zeigen.

## 8. Optional: Windows-Aufgabenplanung

Eine Aufgabe kann nach der erfolgreichen Einrichtung das Skript mit diesen Werten starten:

**Programm:**

```text
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
```

**Argumente:**

```text
-NoLogo -NoProfile -NonInteractive -File "D:\AS400\AS400-GitSync\Sync-AS400SourcesToGit.ps1" -ConfigPath "D:\AS400\AS400-GitSync\AS400-GitSync.psd1" -Execute
```

**Starten in:** `D:\AS400\AS400-GitSync`

Intervall nach Bedarf wählen. Die Aufgabe unter dem Windows-Benutzer ausführen, der die Credential-Datei erstellt hat und auf IFS sowie GitHub zugreifen kann. Gemappte Laufwerke wie `F:` sind in geplanten Tasks häufig nicht vorhanden; dafür den tatsächlichen UNC-Pfad oder einen lokalen Clone konfigurieren. Die bestehende Ausführungsrichtlinie muss das Skript erlauben; das Paket ändert diese Richtlinie nicht.

Es wurde keine geplante Aufgabe auf deinem Rechner angelegt.

## Prüfung des gelieferten Pakets

Die PowerShell-Syntax und 16 Ablaufprüfungen wurden mit PowerShell 7.4.19 und echten lokalen Git-Repositories erfolgreich geprüft. Zusätzlich wurde die Argumentübergabe für den .NET-Framework-Codepfad mit Leerzeichen, Anführungszeichen, Umlauten und abschließenden Backslashes geprüft. IBM-i-Exporte wurden in den Ablaufprüfungen simuliert. Ein Live-Test gegen deine IBM i, Windows PowerShell 5.1, NetServer und GitHub ist hier nicht möglich gewesen. Das Skript verwendet PowerShell-5.1-kompatible Syntax.

## Herstellerdokumentation

- [IBM: QSYS2.QCMDEXC](https://www.ibm.com/support/pages/qsys2qcmdexc)
- [IBM: QCMDEXC-Prozedur, IBM i 7.4](https://www.ibm.com/docs/en/i/7.4.0?topic=services-qcmdexc-procedure)
- [IBM: CPYTOSTMF](https://www.ibm.com/docs/en/i/7.5.0?topic=c-copy-stream-file)
- [IBM: ACS Windows Application Package](https://www.ibm.com/support/pages/ibm-i-access-acs-windows-information)
- [Microsoft: Export-Clixml](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/export-clixml?view=powershell-5.1)
- [Git: git add](https://git-scm.com/docs/git-add)
- [Git: gitattributes und eol](https://git-scm.com/docs/gitattributes)
