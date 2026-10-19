# Kunstløpassistenten

Et program for kunstløpstevner med to deler:

- **Tidsplan/planlegging**: Hent deltakerlisten fra isonen, eller legg inn utøvere selv for mindre stevner. Sett starttid, lengde på løp, klart til neste løp, oppvarming og vanningspauser. Dra rader med musen, så oppdateres tidene med en gang. Oppsettet lagres i en egen fil, og startlisten eksporteres til Excel og PDF.
- **Stevne**: Stor klokke, musikk fra musikktjenesten som spilles av i programmet eller i VLC, og avkrysning av løp som er gjennomført.

Til et stevne trengs tre filer (eller velg «Skann mappe» på mappen der de ligger):

- `Deltakerliste*.xlsx` fra isonen.
- `entries_*.xlsx`, deltakerlisten fra musikktjenesten. Kobles til isonen med NIF ID (= Person ID).
- `Music*.zip` fra musikktjenesten, med lydfiler per klasse og program (`SP`/`FS`) og `music-status.csv`.

Utøvere med kortprogram og friløp får ett løp for hvert program.

Start med `start_kunstlopassistenten.bat`.
<img width="1919" height="1031" alt="Skjermbilde 2026-02-10 194624" src="https://github.com/user-attachments/assets/2aa29d29-f6f6-470f-aaaa-21b31a61caf2" />
