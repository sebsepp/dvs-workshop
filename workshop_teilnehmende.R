# =============================================================================
#
#                          dvs Methoden-Selbsthilfe 🛟
#                                   17.11.2026
#                     Datenaufbereitung & Datenvisualisierung
#
# =============================================================================

# So arbeitest du mit diesem Skript ----
#
#   Die Pipeline:   1 Projekt  ▸  2 Import  ▸  3 Inspizieren  ▸  4 Aufbereiten
#                   ▸  5 Umformen  ▸  6 Visualisieren  ▸  7 Anpassen
#                   ▸  8 Exportieren
#
#   Symbole:
#     👀  Demo           wird live gezeigt - hier steht Hintergrundwissen
#     ✏️  Arbeitsphase   hier bist DU dran
#     💡  Tipp           ein Hinweis, wenn du nicht weiterkommst
#     🛟  Sicherheitsnetz Zwischenstand laden, falls du hängen bleibst
#     🆘  Hängt es?      Fehlermeldung in den Chat schreiben
#
#   Gerüste und Lücken:
#     Code mit ___ ist ein Gerüst: Ersetze jedes ___ durch den passenden Code.
#     Gerüste sind auskommentiert. Markieren und Strg/Cmd + Shift + C
#     entfernt die #-Zeichen.
#
#   Sicherheitsnetz:
#     Wenn ein Schritt nicht klappt, ist das kein Problem! Entferne das #
#     vor der 🛟-Zeile am Anfang des nächsten Blocks, dann hast du den
#     richtigen Zwischenstand und kannst weitermachen.
#
#   Beispieldaten: data/raw/sportstudie.csv  und  data/raw/sportstudie.xlsx
#                  (fiktiv - alle Personen und Werte sind erfunden)

# =============================================================================
#  Pipeline   ● ○ ○ ○ ○ ○ ○ ○        Schritt 1/8: Projekt
# =============================================================================

# 1 | Projekt aufsetzen + Pakete laden ----

## 👀 Pakete ----

# falls noch nicht geschehen ⬇️
# install.packages(c("tidyverse", "readxl", "here", "writexl"))

library(tidyverse) # dplyr, tidyr, ggplot2, readr, lubridate, ...
library(readxl) # Excel-Dateien einlesen
library(here) # robuste Dateipfade
library(writexl) # Excel-Dateien schreiben


## 👀 Reproduzierbar arbeiten ----

# Die 6 Faustregeln:
#   1. Ein Projekt pro Studie (RStudio: File > New Project; Positron: Ordner öffnen)
#   2. Feste Ordnerstruktur: data/raw, data/processed, scripts, output
#   3. Rohdaten NIE überschreiben (nur lesen!)
#   4. Skript statt Klicken: Jeder Schritt steht im Code
#   5. library() immer am Anfang des Skripts
#   6. Workspace NICHT speichern
#      (RStudio: Tools > Global Options > General > Save workspace: "Never")

# here() findet den Projektordner automatisch - egal, wo das Skript liegt
here()
here("data", "raw", "sportstudie.csv")

# Ordnerstruktur per Code anlegen (passiert nichts, wenn sie schon existiert)
c("data/raw", "data/processed", "scripts", "output") |>
  walk(\(ordner) {
    dir.create(here(ordner), recursive = TRUE, showWarnings = FALSE)
  })

# Liegt die Datei dort, wo wir sie erwarten?
file.exists(here("data", "raw", "sportstudie.csv"))

# =============================================================================
#  Pipeline   ○ ● ○ ○ ○ ○ ○ ○        Schritt 2/8: Import
# =============================================================================

# 2 | Daten einlesen ----

## 👀 Welche Funktion für welches Format? ----
#
#   Format                   Funktion                         Paket
#   -----------------------  -------------------------------  ---------
#   .csv (Komma, Punkt)      read_csv()                       readr
#   .csv (Semikolon, Komma)  read_csv2()                      readr
#   .csv/.txt/.tsv (andere)  read_delim(delim = "...")        readr
#   .xlsx / .xls             read_excel()                     readxl
#   SPSS / Stata / SAS       read_sav() / read_dta() / ...    haven
#
#   Der Import-Assistent (File > Import Dataset) hilft beim Ausprobieren.
#   Wichtig: Den erzeugten Code ins Skript kopieren!

## 👀 Checkliste: Worauf achten? ----
#
#   [ ] Trennzeichen?          Komma, Semikolon, Tab ...
#   [ ] Dezimalzeichen?        Punkt oder Komma
#   [ ] Encoding?              UTF-8 oder Windows-1252 (Umlaute kaputt?)
#   [ ] Kopfzeilen?            skip = ...
#   [ ] Fehlcodes?             na = c("", "NA", "-99", "k.A.")
#   [ ] Excel: Sheet / Range?  sheet = ..., range = ...
#   [ ] Spaltentypen korrekt?  Zahlen als Zahl, Datum als Datum?

## 👀 Demo: CSV mit der falschen Funktion ----

csv_pfad <- here("data", "raw", "sportstudie.csv")

falsch <- read_csv(csv_pfad)
glimpse(falsch)
# → Eine einzige Riesenspalte: Das Trennzeichen ist ein Semikolon!

## 👀 Demo: Trennzeichen richtig, Dezimalzeichen noch nicht ----

halb_richtig <- read_delim(csv_pfad, delim = ";", show_col_types = FALSE)
glimpse(halb_richtig)
# → Die Spalten sind jetzt getrennt, ABER: Das Komma gilt standardmäßig als
#   Tausendertrenner. Aus 80,1 kg wird 801 - ohne Fehlermeldung! 😱
#   → Immer prüfen, ob die Werte plausibel sind, nicht nur die Spaltentypen.

## 👀 Demo: CSV richtig einlesen ----

#
# Gerüst:
# sportstudie_raw <- ___(
#   csv_pfad,
#   na = c(___),            # welche Fehlcodes gibt es? (siehe Checkliste)
#   show_col_types = FALSE
# )
# glimpse(sportstudie_raw)

# ⚠️ Achtung: na = ... gilt für ALLE Spalten. Das ist nur okay, wenn z. B. -99
#    nirgends ein echter Wert ist (Temperatur in °C wäre es!).

# Probleme mit Umlauten (z. B. "MÃ¤nnlich")? Dann stimmt das Encoding:
# guess_encoding(csv_pfad)
# read_csv2(csv_pfad, locale = locale(encoding = "Windows-1252"))

## 👀 Demo: Excel einlesen ----

xlsx_pfad <- here("data", "raw", "sportstudie.xlsx")

excel_sheets(xlsx_pfad) # Welche Blätter gibt es?

# Zweites Blatt: Codebook (immer lesen, wenn es eins gibt!)
read_excel(xlsx_pfad, sheet = "Codebook")

# Gerüst für das Datenblatt:
# sportstudie_raw <- read_excel(
#   xlsx_pfad,
#   sheet = "___",
#   skip  = ___,            # 💡 In welcher Zeile beginnt die Tabelle? (Datei öffnen!)
#   na    = c(___)
# )

# Nur einen Ausschnitt lesen: read_excel(..., range = "A3:F20")

# Weitere Formate:
# library(haven)
# read_sav("daten.sav")   # SPSS
# read_dta("daten.dta")   # Stata

# =============================================================================
#  Pipeline   ○ ○ ● ○ ○ ○ ○ ○        Schritt 3/8: Inspizieren
# =============================================================================

# 3 | Daten verstehen ----

# 🛟 Sicherheitsnetz: Import hat nicht geklappt? Dann das # entfernen ⬇️
# sportstudie_raw <- readRDS(here("data", "checkpoints", "01_sportstudie_raw.rds"))

## 👀 Erst anschauen, dann bereinigen ----

glimpse(sportstudie_raw) # Struktur: Spalten, Typen, erste Werte
summary(sportstudie_raw) # Min/Max/Mittelwert + Anzahl NA je Spalte
names(sportstudie_raw) # Spaltennamen
View(sportstudie_raw) # Tabelle ansehen (RStudio: View, Positron: Data Explorer)

# Kategoriale Variablen: Welche Ausprägungen gibt es?
sportstudie_raw |> count(Geschlecht)
sportstudie_raw |> count(Erfahrung)
sportstudie_raw |> count(Sportart, Gruppe)

# Fehlende Werte pro Spalte
colSums(is.na(sportstudie_raw))

# Dubletten?
sum(duplicated(sportstudie_raw))
sportstudie_raw |> count(ID) |> filter(n > 1)

# Auffällige Werte?
sportstudie_raw |> arrange(desc(Alter)) |> select(ID, Alter) |> head()

# Optional: ein Rundum-Überblick
# install.packages("skimr")
# skimr::skim(sportstudie_raw)

# =============================================================================
#  Pipeline   ○ ○ ○ ● ○ ○ ○ ○        Schritt 4/8: Aufbereiten
# =============================================================================

# 4 | Daten aufbereiten mit dplyr ----

## 👀 Die wichtigsten dplyr-Verben ----
#
#   select()      Spalten auswählen
#   filter()      Zeilen auswählen
#   mutate()      neue Variablen erzeugen / verändern
#   arrange()     sortieren
#   rename()      umbenennen (neu = alt)
#   distinct()    Dubletten entfernen
#   count()       Häufigkeiten zählen
#   group_by() + summarise()   Kennwerte je Gruppe

sportstudie_raw |> select(ID, Gruppe, Alter) # Spalten wählen
sportstudie_raw |> select(starts_with("VO2max")) # nach Muster wählen
sportstudie_raw |> filter(Gruppe == "Intervention") # Zeilen wählen
sportstudie_raw |> filter(Alter > 30, Sportart == "Kraft") # mehrere Bedingungen (UND)
sportstudie_raw |> arrange(desc(`VO2max T1`)) # absteigend sortieren

## 👀 Die Pipe |> ("und dann") ----

# Ohne Pipe: verschachtelt, von innen nach außen lesen
head(arrange(filter(sportstudie_raw, Gruppe == "Kontrolle"), Alter), 3)

# Mit Pipe: von oben nach unten lesen - "nimm die Daten, UND DANN ..."
sportstudie_raw |>
  filter(Gruppe == "Kontrolle") |>
  arrange(Alter) |>
  head(3)

# Hinweis: In älteren Skripten / Tutorials steht oft %>% (magrittr-Pipe).
# Für unsere Zwecke funktioniert sie praktisch genauso.

## 👀 Werkzeugkasten: neue Variablen und Kategorien ----

# Allgemeine Muster (mit Fantasie-Variablen - NICHT unser Datensatz):
#
#   mutate(neu = a / b)                              # berechnen
#   mutate(neu = if_else(x > 10, "hoch", "niedrig")) # zwei Kategorien
#   mutate(neu = case_when(                          # mehrere Kategorien
#     x < 10  ~ "klein",                             #   (die Reihenfolge der
#     x < 20  ~ "mittel",                            #    Bedingungen zählt!)
#     x >= 20 ~ "groß"
#   ))
#   mutate(neu = if_else(x > 100, NA, x))            # Werte auf NA setzen
#   mutate(x = factor(x, levels = c("a", "b", "c"))) # ordinal: feste Reihenfolge
#   mutate(datum = dmy("29.09.2025"))                # Text → Datum (lubridate)
#   mutate(x = x == "ja")                            # Text → TRUE/FALSE
#
#   Neue Variable prüfen: count(), glimpse(), View() - nach JEDEM Schritt!

## 👀 Live-Demo: Daten aufräumen (gemeinsam) ----

# Regel für Spaltennamen: klein, keine Leerzeichen, keine Umlaute, keine
# Sonderzeichen. Mit rename(NEU = ALT). Namen mit Sonderzeichen in Backticks ``.

# Gerüst Schritt 1:
# sportstudie <- sportstudie_raw |>
#   rename(
#     id         = ID,
#     gruppe     = ___,
#     groesse_cm = `Größe (cm)`,
#     gewicht_kg = ___,
#     ...                      # ← alle 16 Spalten umbenennen
#     vo2max_t3  = ___
#   )

# Gerüst Schritt 2 (Rohformat: "29.09.2025" = Tag.Monat.Jahr):
# sportstudie <- sportstudie |>
#   mutate(erhebungsdatum = ___(erhebungsdatum))

# Gerüst Schritt 3 (erst count(geschlecht) ansehen!):
# sportstudie <- sportstudie |>
#   mutate(
#     geschlecht = case_when(
#       str_to_lower(geschlecht) %in% c(___) ~ "weiblich",
#       str_to_lower(geschlecht) %in% c(___) ~ "männlich",
#       .default = ___                       # alles andere unverändert lassen
#     )
#   )

# Gerüst Schritt 4 (erst filter(alter > 60) ansehen!):
# sportstudie <- sportstudie |>
#   mutate(alter = if_else(___, NA, alter))
# Wichtig: Entscheidung dokumentieren, Rohdaten NICHT ändern!

# Gerüst Schritt 5:
# sportstudie <- sportstudie |>
#   mutate(bmi = ___ / (___ / 100)^2)

# 🛟 Sicherheitsnetz: Aufbereitung nicht fertig? Dann das # entfernen ⬇️
# sportstudie <- readRDS(here("data", "checkpoints", "02_sportstudie_clean.rds"))

## 👀 Gruppieren und zusammenfassen ----

# group_by() + summarise(): ein Kennwert-Set pro Gruppe
sportstudie |>
  group_by(gruppe) |>
  summarise(
    n = n(),
    mw_vo2max = mean(vo2max_t1, na.rm = TRUE), # na.rm = NA ignorieren
    sd_vo2max = sd(vo2max_t1, na.rm = TRUE)
  )

# Mehrere Gruppierungsvariablen: group_by(gruppe, sportart)
# Danach die Gruppierung auflösen: summarise(..., .groups = "drop")
# Mehrere Spalten auf einmal: summarise(across(starts_with("vo2max"), ...))

# =============================================================================
#  Pipeline   ○ ○ ○ ○ ● ○ ○ ○        Schritt 5/8: Umformen
# =============================================================================

# 5 | Long vs. Wide ----

## 👀 Warum umformen? ----
#
#   WIDE: eine Zeile pro Person, Messzeitpunkte in eigenen Spalten
#         id | vo2max_t1 | vo2max_t2 | vo2max_t3
#
#   LONG: eine Zeile pro Messung
#         id | zeitpunkt | vo2max
#
#   Faustregel für ggplot2:
#   Alles, was auf eine Achse, eine Farbe oder in eine Facette soll,
#   braucht eine EIGENE SPALTE → also LONG.

## 👀 pivot_longer() und pivot_wider() ----

# pivot_longer(): wide → long
#   Argumente:
#     cols         Welche Spalten sollen zusammengeführt werden?
#     names_to     In welche NEUE Spalte kommen die alten Spaltennamen?
#     names_prefix Welcher Teil des Namens soll abgeschnitten werden?
#     values_to    In welche NEUE Spalte kommen die Werte?
#
# pivot_wider(): long → wide
#   Argumente:
#     names_from   Aus welcher Spalte kommen die neuen Spaltennamen?
#     values_from  Aus welcher Spalte kommen die Werte?
#
# 💡 Tidyselect-Helfer für cols: starts_with("..."), ends_with("..."), c(a, b, c)

# Gerüst pivot_longer():
# vo2_long <- sportstudie |>
#   select(id, gruppe, geschlecht, sportart, erfahrung, starts_with("vo2max")) |>
#   pivot_longer(
#     cols         = ___("vo2max"),
#     names_to     = "___",
#     names_prefix = "___",
#     values_to    = "___"
#   ) |>
#   mutate(zeitpunkt = factor(zeitpunkt, levels = c(___), labels = c("T1", "T2", "T3")))

# Danach prüfen: glimpse(vo2_long); nrow(sportstudie); nrow(vo2_long)
# Erwartung: 3 × so viele Zeilen wie vorher. Stimmt das bei dir?

# =============================================================================
#  Pipeline   ○ ○ ○ ○ ○ ● ○ ○        Schritt 6/8: Visualisieren
# =============================================================================

# 6 | Visualisieren mit ggplot2 ----

# 🛟 Sicherheitsnetz: Fehlt etwas? Dann das # entfernen ⬇️
# sportstudie <- readRDS(here("data", "checkpoints", "02_sportstudie_clean.rds"))
# vo2_long    <- readRDS(here("data", "checkpoints", "03_vo2_long.rds"))

## 👀 Die Grammatik einer Grafik ----
#
#   ggplot(data = ..., mapping = aes(x = ..., y = ..., color/fill = ...)) +
#     geom_...()      ← Welche Form? (Punkte, Balken, Boxen ...)
#     facet_...()     ← Aufteilen in Teilgrafiken
#     scale_...()     ← Achsen und Farben anpassen
#     labs()          ← Beschriftungen
#     theme_...()     ← Gesamtdesign
#
#   Eine Grafik entsteht Schicht für Schicht, verbunden mit  +
#   (Achtung: In ggplot2 steht ein "+", nicht die Pipe, am Zeilenende!)
#
#   aes() = Variable → visuelle Eigenschaft ("mapping")
#   außerhalb von aes() = fester Wert für alle (z. B. color = "red")

## 👀 Welcher Plot für welches Skalenniveau? ----
#
#   Variable(n)                        Geom
#   ---------------------------------  ------------------------------------
#   nominal / ordinal (Häufigkeiten)   geom_bar() / geom_col()
#   metrisch, eine Variable            geom_histogram() / geom_density()
#   metrisch nach Gruppe               geom_boxplot() / geom_violin() + geom_jitter()
#   metrisch × metrisch                geom_point() + geom_smooth()
#   Verlauf über die Zeit              geom_line() + stat_summary()

### nominal: Häufigkeiten ----
ggplot(sportstudie, aes(x = sportart)) +
  geom_bar()

# ... aufgeteilt nach Gruppe (nebeneinander)
ggplot(sportstudie, aes(x = sportart, fill = gruppe)) +
  geom_bar(position = "dodge")

### metrisch (1 Variable): Verteilung ----
ggplot(sportstudie, aes(x = trainingsumfang_h)) +
  geom_histogram(bins = 20) # rechtsschief: wenige haben sehr viel Umfang

ggplot(sportstudie, aes(x = trainingsumfang_h, fill = sportart)) +
  geom_density(alpha = 0.4) # alpha = Transparenz

### metrisch nach Gruppe: Violin + Rohdaten ----
ggplot(sportstudie, aes(x = sportart, y = vo2max_t1)) +
  geom_violin() +
  geom_jitter(width = 0.1, alpha = 0.5) # Rohdaten als Punkte mit leichtem Versatz

### metrisch × metrisch: Streudiagramm ----
ggplot(sportstudie, aes(x = alter, y = vo2max_t1)) +
  geom_point() +
  geom_smooth(method = "lm") # lineare Regressionsgerade mit Konfidenzband

# Hinweis zur Warnung "Removed ... rows containing missing values":
# ggplot2 lässt fehlende Werte weg und sagt uns Bescheid (hier: das Alter, das
# wir auf NA gesetzt haben). Mit na.rm = TRUE im Geom wird die Warnung unterdrückt.

## 👀 Live-Aufbau: Schicht für Schicht ----

# Wir bauen gemeinsam einen Plot der VO2max über die Zeit auf:
#
#   1) Basis speichern:       p <- ggplot(___, aes(x = ___, y = ___))
#   2) Boxplot:               p + geom_boxplot(na.rm = TRUE)
#   3) Farbe nach Gruppe:     aes(fill = ___) im Geom
#   4) Facetten:              + facet_wrap(~ ___)
#   5) Mittelwerte:           + stat_summary(fun = mean, geom = "point", ...)
#
# Alternative: Mittelwert ± Standardfehler als Linie
#   stat_summary(fun = mean, geom = "line")
#   stat_summary(fun.data = mean_se, geom = "errorbar", width = 0.15)
#   (dafür braucht ggplot2 group = ... und color = ... in aes())
#
# 💡 Welche Daten (wide oder long?) brauchst du dafür?

# =============================================================================
#  Pipeline   ○ ○ ○ ○ ○ ○ ● ○        Schritt 7/8: Anpassen
# =============================================================================

# 7 | Grafik gestalten ----

## 👀 Werkzeugkasten: Beschriftung, Skalen, Farben, Themes ----
#
#   Beschriftung    labs(title = , subtitle = , x = , y = , color = , caption = )
#   Farben          scale_color_manual(values = ...)   scale_fill_manual(values = ...)
#                   scale_color_viridis_d()            scale_fill_viridis_d()
#                   → farbenblind-freundlich!
#   Achsen          scale_y_continuous(breaks = ..., limits = ...)  schneidet Daten ab
#                   coord_cartesian(ylim = c(..., ...))             zoomt nur hinein
#   Fertige Themes  theme_minimal()  theme_classic()  theme_bw()   (base_size = ...)
#   Feinjustierung  theme(legend.position = "bottom",
#                         plot.title = element_text(face = "bold"),
#                         panel.grid.minor = element_blank(),
#                         strip.text = element_text(face = "bold"))
#   Für alle Plots  theme_set(theme_minimal(base_size = 12))

# Eigene Farben (Okabe-Ito-Palette, farbenblind-freundlich)
farben <- c("Intervention" = "#0072B2", "Kontrolle" = "#E69F00")

# Gerüst - `deine_grafik` ist die Grafik aus Arbeitsphase 4:
# p_final <- deine_grafik +
#   labs(
#     title    = "___",
#     x        = "___",
#     y        = "___ (Einheit!)",
#     color    = "___"                  # oder fill = ... (je nach Plot)
#   ) +
#   scale_color_manual(values = farben) +   # oder scale_fill_manual()
#   theme_<name>(base_size = ___) +       # 💡 <name> = minimal, classic oder bw
#   theme(legend.position = "___")
# p_final

# =============================================================================
#  Pipeline   ○ ○ ○ ○ ○ ○ ○ ●        Schritt 8/8: Exportieren
# =============================================================================

# 8 | Ergebnisse exportieren ----

## 👀 Grafiken speichern: ggsave() ----

# Gerüst:
# ggsave(
#   filename = here("output", "___.png"),
#   plot     = ___,                       # ⚠️
#   width    = ___, height = ___, units = "cm",
#   dpi      = ___
# )
#
# 💡 Orientierung: 1 Spalte ≈ 8-9 cm, 2 Spalten ≈ 17-18 cm Breite (Journal prüfen!)
# 💡 Die Dateiendung bestimmt das Format: .png, .pdf, .tiff, ...

# Format-Faustregel:
#   PNG  → Folien, Word, Web
#   PDF / SVG → Vektorgrafik für Publikationen (SVG braucht das Paket svglite)
#   TIFF → manche Journals verlangen es (dpi = 300-600)

# ⚠️ Ohne plot = ... wird die zuletzt angezeigte Grafik gespeichert.

## 👀 Daten speichern ----

# Bereinigte Daten gehören nach data/processed (NIE nach data/raw!)
#
#   Funktion                         Format                          Paket
#   -------------------------------  ------------------------------  ---------
#   write_csv2(daten, pfad)          CSV deutsch (; und ,)           readr
#   write_csv(daten, pfad)           CSV international (, und .)     readr
#   write_xlsx(daten, pfad)          Excel                           writexl
#   saveRDS(daten, pfad)             R-Format (behält Faktoren)      base R
#   readRDS(pfad)                    R-Format wieder einlesen        base R
#
# Gerüst:
# write_csv2(___, here("data", "processed", "___.csv"))

# Kontrolle: Liegt alles dort, wo es hingehört?
list.files(here("output"))
list.files(here("data", "processed"))

# =============================================================================
#  Abschluss
# =============================================================================

# 9 | Hilfe zur Selbsthilfe ----

## 👀 Hilfe in R ----

?mutate # Hilfeseite einer Funktion
??"pivot" # Stichwortsuche in allen Hilfeseiten
vignette("dplyr") # ausführliche Anleitung zu einem Paket
example(geom_boxplot) # Beispiele einer Funktion ausführen


## 👀 Fehlermeldungen lesen ----
#
#   Fehlermeldung                         Meistens bedeutet das ...
#   ------------------------------------  --------------------------------------
#   object 'xyz' not found                Tippfehler ODER Objekt nicht erstellt
#                                         (Zeile nicht ausgeführt?)
#   could not find function "xyz"         Paket nicht geladen (library()?)
#                                         oder Tippfehler im Funktionsnamen
#   unexpected symbol / unexpected ')'    Klammer, Komma oder Anführungszeichen
#                                         fehlt / zu viel
#   Can't subset columns that don't       Spaltenname existiert nicht (Tippfehler?
#   exist                                 Umbenannt? Groß-/Kleinschreibung?)
#   non-numeric argument to binary        Du rechnest mit Text (Spalte ist
#   operator                              character, nicht numeric → Import prüfen!)
#
#   Vorgehen: 1) Fehlermeldung LESEN  2) Zeilennummer ansehen
#             3) Code Stück für Stück ausführen  4) Googeln (Meldung kopieren)

## 👀 Weiterführende Quellen ----
#
#   Cheatsheets (Posit):   https://posit.co/resources/cheatsheets/
#   R for Data Science:    https://r4ds.hadley.nz/
#   tidyverse-Website:     https://www.tidyverse.org/
#   ggplot2-Galerie:       https://r-graph-gallery.com/
#   Community:             https://forum.posit.co/
#
#   Fragen stellen? → mit minimalem, reproduzierbarem Beispiel (reprex):
#   # install.packages("reprex")
#   # reprex::reprex()
#
#   🤖 KI-Assistenten (ChatGPT, Claude, Copilot ...) können beim Coden helfen.
#      ABER: Code immer selbst prüfen und verstehen, und NIE echte, personen-
#      bezogene (z. B. sportmedizinische) Daten in Chatfenster kopieren!

sessionInfo() # Welche R- und Paketversionen? (hilfreich bei Fehlern + Methodenteil)

# Ausblick ----

#   Weiterlernen:  Joins (left_join), Quarto-Berichte, Tabellen (gt, flextable),
#                  mehrere Plots kombinieren (patchwork), Statistik (rstatix, lme4)
#   Material:      GitHub-Repository zum Workshop (Link folgt)
#                  Das vollständige Skript mit allen Lösungen folgt nach dem Workshop.
