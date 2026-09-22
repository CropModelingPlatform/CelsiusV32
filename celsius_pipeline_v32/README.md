# Pipeline CELSIUS VBA -> VB .NET

Ce dossier formalise un flux de synchronisation a sens unique:

- `Access/VBA` reste la source de verite pour le code scientifique.
- `VB .NET` est regenere a partir des exports Access.
- Le code `.NET` ecrit a la main reste hors de `Converted/`.

Le point important est de ne pas faire de synchronisation bidirectionnelle. Les modules VBA exportes alimentent `Converted/`, et les fichiers hors `Converted/` restent stables et maintenus cote .NET.

## Structure du flux

1. Export des objets Access en texte.
2. Conversion des modules VBA exportes en fichiers `.vb`.
3. Synchronisation du projet VB .NET cible.
4. Export des tables d'entree Access.
5. Construction d'une base SQLite d'entree.
6. Build et execution du CLI VB .NET.
7. Test de regression en comparant les sorties Access et SQLite.

## Fichiers du pipeline

- `export_access_objects.ps1`
  - Exporte modules, formulaires, rapports et requetes depuis `.accdb`.
- `export_access_tables.ps1`
  - Exporte des tables Access en TSV.
- `convert_vba_to_vbnet.py`
  - Regenerer `Converted/*.vb` depuis les `.bas` exportes.
- `build_sqlite_from_tsv.py`
  - Construit la base SQLite d'entree a partir des TSV.
- `sync_celsius_model.ps1`
  - Script principal de synchronisation.
- `run_access_model.ps1`
  - Lance le modele VBA dans Access.
- `export_sqlite_tables.py`
  - Exporte les tables SQLite de sortie en TSV.
- `compare_tsv_tables.py`
  - Compare les sorties Access et SQLite avec tolerance numerique.
- `regression_test.ps1`
  - Lance un test de regression complet.

## Hypotheses

- La base Access du collegue est dans `/path/celsiusdb.accdb`.
- Le projet VB .NET cible doit etre genere dans `/path2/celsiusCli`.
- Le squelette manuel du projet VB .NET est le projet present dans `../CelsiusCli`.
- `Converted/` est entierement regenere a chaque sync.

## Etapes de transformation

### 1. Export Access -> texte

`export_access_objects.ps1` cree un export texte du code et des objets Access:

- modules `.bas`
- formulaires `.form.txt`
- rapports `.report.txt`
- requetes `.query.txt`

Exemple:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\celsius_pipeline\export_access_objects.ps1 `
  -AccessPath /path/celsiusdb.accdb `
  -ExportDir .\work\vba_export
```

### 2. Conversion texte VBA -> VB .NET genere

`convert_vba_to_vbnet.py` lit les `.bas` exportes et remplit `Converted/` dans le projet cible:

```powershell
python .\celsius_pipeline\convert_vba_to_vbnet.py `
  .\work\vba_export `
  /path2/celsiusCli
```

Regles actuelles:

- `Functions` reste un `Module`.
- `Principal` et `NumeroteEnreg` sont ignores car l'entree .NET est geree ailleurs.
- `GraphesJournaliers` et `Compteur_param` sont ignores car ce sont des utilitaires Access, pas des composants du moteur scientifique.
- Les fichiers existants dans `Converted/` sont supprimes avant regeneration.

### 3. Export des tables d'entree et creation du SQLite

`export_access_tables.ps1` exporte les tables d'entree en TSV, puis `build_sqlite_from_tsv.py` construit la base SQLite:

Le pipeline exporte aussi automatiquement le schema de la base avec
`export_access_schema.ps1`. Le schema SQLite n'est donc plus lie a une ancienne
version de `CelsiusV32.accdb`. Les types Access `Single` et `Double` restent
distincts afin de conserver le comportement numerique du modele.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\celsius_pipeline\export_access_tables.ps1 `
  -AccessPath /path/celsiusdb.accdb `
  -OutputDir .\work\input_tsv

python .\celsius_pipeline\build_sqlite_from_tsv.py `
  .\celsius_table_schema.csv `
  .\work\input_tsv `
  .\work\celsius_model_input.db
```

### 4. Build et execution du CLI VB .NET

```powershell
dotnet build /path2/celsiusCli/CelsiusCli.vbproj
dotnet run --project /path2/celsiusCli/CelsiusCli.vbproj --no-build -- .\work\celsius_model_input.db
```

## Script de synchronisation

Le script `sync_celsius_model.ps1` enchaine tout le flux de generation:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\celsius_pipeline\sync_celsius_model.ps1 `
  -AccessPath /path/celsiusdb.accdb `
  -TargetProjectDir /path2/celsiusCli `
  -BuildProject
```

Ce script fait:

1. Creation du projet cible a partir du template `../CelsiusCli` s'il n'existe pas encore.
2. Export du code VBA de la base Access.
3. Regeneration de `/path2/celsiusCli/Converted`.
4. Export des tables d'entree.
5. Construction de `work/celsius_model_input.db`.
6. Build optionnel du projet cible.

Sorties de travail:

- `celsius_pipeline/work/vba_export`
- `celsius_pipeline/work/input_tsv`
- `celsius_pipeline/work/celsius_model_input.db`

## Installation comme commande Windows

Le projet VB .NET peut etre installe comme commande systeme, par exemple `celsius`.

### Ce que fait l'installation

1. `dotnet publish` du projet cible.
2. Creation d'un lanceur `celsius.cmd`.
3. Ajout optionnel du dossier d'installation au `PATH` utilisateur.

### Commande d'installation

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\celsius_pipeline\install_windows_command.ps1 `
  -ProjectDir /path2/celsiusCli `
  -AddToUserPath
```

Par defaut, le script installe dans `%LOCALAPPDATA%\CelsiusCli` et cree la commande `celsius`.

### Utilisation

```powershell
celsius --help
celsius C:\data\celsius_model_input.db
```

## Test de regression

Le test de regression est volontairement separe du script de sync. Il suppose que le projet VB .NET cible est deja synchro.

### Principe

1. Copier la base Access source dans un dossier de travail.
2. Lancer le modele VBA dans cette copie.
3. Exporter les tables d'entree et de sortie Access.
4. Reconstruire une base SQLite d'entree identique.
5. Executer le CLI VB .NET sur cette base SQLite.
6. Exporter `OutputSynt`, `OutputD_1`, `OutputD_2`.
7. Comparer les fichiers TSV Access et SQLite.

### Commande

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\celsius_pipeline\regression_test.ps1 `
  -AccessPath /path/celsiusdb.accdb `
  -ProjectDir /path2/celsiusCli
```

### Entree VBA utilisee

Par defaut, le test appelle `Zyva`, qui appelle ensuite `Principal` dans le code VBA exporte. Si votre base Access a besoin d'une autre entree, utilisez `-EntryPoint`.

Exemple:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\celsius_pipeline\regression_test.ps1 `
  -AccessPath /path/celsiusdb.accdb `
  -ProjectDir /path2/celsiusCli `
  -EntryPoint MyBatchLauncher
```

## Gouvernance recommandee

- Le collegue qui maintient Access modifie le code dans `.accdb`.
- On exporte le code Access en texte et on versionne ces exports.
- On regenere `Converted/` automatiquement.
- Les fichiers manuels du projet `.NET` ne doivent pas etre modifies par le pipeline.
- On interdit les editions manuelles dans `Converted/`.

## Dossier cible VB .NET

Dans le projet cible `/path2/celsiusCli`:

- `Converted/`
  - code genere depuis VBA
- tout le reste
  - code manuel `.NET` pour CLI, SQLite, runtime et compatibilite

## Limitations actuelles

- Le pipeline suppose Microsoft Access installe et pilotable par COM sous Windows.
- Le test de regression suppose que l'entree VBA peut etre lancee sans intervention utilisateur via `Application.Run`.
- Une dependance historique a `ListResidus` existe dans `ApportsOrgaClass`; si cette table n'est pas presente dans la base, le chemin correspondant reste un point de vigilance.
- Les tables `CO2Yearly` et `ListResidus` font partie des entrees du moteur V32.
- La regression ignore explicitement `BilanNnonOK` et `DrainageON`: ces deux indicateurs changent au voisinage exact d'un seuil flottant. Les valeurs continues dont ils derivent restent comparees.
- La comparaison des sorties est un test de regression pragmatique base sur des exports TSV avec tolerance numerique, pas une preuve formelle d'equivalence.

## Workflow recommande au quotidien

1. Le collegue modifie la base Access.
2. Lancer `sync_celsius_model.ps1`.
3. Construire le projet cible.
4. Lancer `regression_test.ps1`.
5. Si la regression passe, valider les changements cote `.NET`.
