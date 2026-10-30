# Kunstløpassistenten

Et program for kunstløpstevner. Det leser deltakerlister fra **isonen** og musikk fra **NSF sin musikkside ([nsfcompetition.no](https://nsfcompetition.no/))**, lager tidsskjema og startliste, og spiller musikken i riktig rekkefølge under stevnet.

## Last ned

**[⬇ Last ned Kunstlopassistenten.zip](https://github.com/Jollymix/KunstlopFikser/releases/latest/download/Kunstlopassistenten.zip)** (for Windows, ca. 40 MB)

Du trenger ikke installere noe annet.

1. Klikk på lenken over. Filen havner i mappen **Nedlastinger**.
2. Høyreklikk på `Kunstlopassistenten.zip` og velg **Pakk ut alle…**, så **Pakk ut**.
   Ikke start programmet direkte fra zip-filen, da virker det ikke.
3. Åpne mappen `Kunstlopassistenten` som kom ut, og dobbeltklikk på **Kunstlopassistenten.exe**.

**Kommer det en blå boks med «Windows beskyttet PCen»?** Det er vanlig for små programmer som ikke er kjøpt i en butikk. Klikk **Mer info** og så **Kjør likevel**. Dette må du bare gjøre første gang.

**Vil du ha programmet på en USB-pinne?** Kopier hele mappen `Kunstlopassistenten` (ikke bare .exe-filen) over til pinnen. Da kan du starte det fra pinnen på hvilken som helst Windows-PC.

![Tidsplan/planlegging](docs/skjermbilde_tidsplan.png)

![Stevne](docs/skjermbilde_stevne.png)

## Hva kan programmet brukes til?

- **Interne stevner og prøvestevner:** Lag et tidsskjema på noen minutter. Hent utøverne fra isonen eller Spond, eller skriv dem inn selv. Velg starttid, lengde på løp, oppvarmingsgrupper og vanningspauser, så regner programmet ut alle tidene. Dra utøverne med musen for å endre rekkefølgen. Startlisten lagres som Excel og PDF.
- **Stevner med offisielt tidsskjema fra Norges Skøyteforbund:** Hent utøverne og musikken fra isonen og nsfcompetition.no, og sett dem i samme rekkefølge som tidsskjemaet fra forbundet. Da ligger musikken klar i riktig rekkefølge når stevnet starter.
- **Under stevnet:** Stor klokke, musikken spilles rett fra programmet (eller i VLC), og du krysser av hvert løp som er gjennomført.

## Filer programmet kan lese

Alt hentes fra **Fil**-menyen i programmet.

| Hvor filen kommer fra | Fil | Hva den brukes til |
|---|---|---|
| **isonen** | `Deltakerliste….xlsx` | Utøverne som er påmeldt stevnet |
| **NSF musikk** ([nsfcompetition.no](https://nsfcompetition.no/)) | `entries_….xlsx` | Deltakerlisten med program (kortprogram/friløp) |
| **NSF musikk** ([nsfcompetition.no](https://nsfcompetition.no/)) | `Music….zip` | Musikken til alle utøverne |
| **Spond** (prøvestevner) | eksportert deltakerliste (`.xlsx`) | Utøverne som har svart ja |
| **En mappe på PC-en** (prøvestevner) | lydfiler, f.eks. `Fornavn_Etternavn_FS.mp3` | Musikken til utøverne |

Utøverne fra isonen og NSF musikk kobles sammen med NIF ID, og musikken kobles til riktig utøver automatisk. Utøvere med både kortprogram og friløp får ett løp for hvert program.

## For utviklere

Start fra kildekoden med `start_kunstlopassistenten.bat` (krever Python og `pip install -r requirements.txt`).

Ny nedlastbar versjon lages automatisk når en versjonstagg pushes: `git tag v1.1; git push --tags`. For å bygge lokalt: `powershell -ExecutionPolicy Bypass -File build_usb.ps1`, resultatet havner i `dist\Kunstlopassistenten`.
