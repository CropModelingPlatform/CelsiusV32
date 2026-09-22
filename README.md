# CELSIUS V32 CLI

Ce dépôt contient le portage en ligne de commande VB.NET du modèle de culture
CELSIUS V32, les sources VBA exportées depuis Microsoft Access et les outils qui
permettent de régénérer le portage et de vérifier sa conformité au modèle Access.

Le flux est volontairement à sens unique :

```text
CelsiusV32.accdb (VBA, source scientifique)
        │
        ├── export des objets Access → celsius_audit_new/access_objects
        │
        ├── conversion des modules .bas
        ▼
CelsiusCli_V32_work/Converted/*.vb
        │
        ├── compilation .NET 9
        ▼
CelsiusCli / celsiusV32
```

Les fichiers de `CelsiusCli_V32_work/Converted` sont générés : les corrections
scientifiques durables doivent être faites dans la base Access/VBA, puis
réexportées et reconverties. Le code d'intégration .NET situé hors de
`Converted/` est maintenu directement dans ce dépôt.

## Organisation du dépôt

- `CelsiusCli_V32_work/` : projet VB.NET 9 compilable sous Linux et Windows.
- `CelsiusCli_V32_work/Converted/` : moteur scientifique généré depuis les
  modules VBA V32.
- `celsius_pipeline_v32/` : export Access, conversion VBA → VB.NET, création de
  la base SQLite et tests de régression.
- `celsius_audit_new/access_objects/` : export texte versionné des objets de la
  base Access V32 (`.bas`, formulaires et requêtes).
- `.github/workflows/ci.yml` : compilation et test minimal automatiques sous
  Linux à chaque push et pull request.

La base `CelsiusV32.accdb`, les bases SQLite, les résultats, les répertoires
`bin/`, `obj/`, `publish/` et les espaces de travail de régression sont exclus
de Git.

## Prérequis

Pour compiler et exécuter le CLI :

- SDK .NET 9 ;
- Linux x64 ou Windows x64.

Pour régénérer le projet et effectuer la régression complète :

- Windows ;
- Microsoft Access installé et accessible par COM ;
- PowerShell ;
- Python 3 ;
- SDK .NET 9 ;
- une copie locale de `CelsiusV32.accdb`.

## Compiler le CLI

Depuis la racine du dépôt :

```bash
dotnet restore CelsiusCli_V32_work/CelsiusCli.vbproj --runtime linux-x64
dotnet build CelsiusCli_V32_work/CelsiusCli.vbproj \
  --configuration Release \
  --runtime linux-x64 \
  --no-restore
```

Pour produire un exécutable Linux dépendant du runtime .NET installé :

```bash
dotnet publish CelsiusCli_V32_work/CelsiusCli.vbproj \
  --configuration Release \
  --runtime linux-x64 \
  --self-contained false \
  --output publish/linux-x64
```

Test minimal :

```bash
./publish/linux-x64/CelsiusCli --help
```

## Exécuter une simulation

Le CLI reçoit le chemin d'une base SQLite contenant les entrées CELSIUS :

```bash
dotnet run \
  --project CelsiusCli_V32_work/CelsiusCli.vbproj \
  --configuration Release \
  -- /chemin/vers/celsius_model_input.db
```

Avec la commande installée localement :

```bash
celsiusV32 /chemin/vers/celsius_model_input.db
```

Le programme exécute les simulations de `SimUnitList` et crée ou réutilise les
tables de sortie suivantes :

- `OutputSynt` ;
- `OutputD_1` ;
- `OutputD_2`.

Les tables `CO2Yearly` et `ListResidus` font partie des entrées spécifiques
prises en charge par le pipeline V32.

## Régénérer le portage depuis Access

Sous Windows PowerShell, depuis la racine du dépôt :

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File .\celsius_pipeline_v32\sync_celsius_model.ps1 `
  -AccessPath D:\chemin\vers\CelsiusV32.accdb `
  -TargetProjectDir .\CelsiusCli_V32_work `
  -BuildProject
```

Le script :

1. exporte les objets VBA dans son espace de travail ;
2. régénère `CelsiusCli_V32_work/Converted` ;
3. exporte les tables et le schéma Access ;
4. construit `celsius_pipeline_v32/work/celsius_model_input.db` ;
5. compile le projet lorsque `-BuildProject` est fourni.

Pour actualiser également l'instantané VBA versionné :

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File .\celsius_pipeline_v32\export_access_objects.ps1 `
  -AccessPath D:\chemin\vers\CelsiusV32.accdb `
  -ExportDir .\celsius_audit_new\access_objects
```

Vérifier ensuite que la conversion de cet export ne produit aucune différence
inattendue dans `Converted/`.

## Régression Access contre VB.NET

Le test complet exécute le modèle VBA dans une copie de la base Access, crée une
base SQLite équivalente, exécute le CLI puis compare `OutputSynt`, `OutputD_1`
et `OutputD_2` :

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File .\celsius_pipeline_v32\regression_test.ps1 `
  -AccessPath D:\chemin\vers\CelsiusV32.accdb `
  -ProjectDir .\CelsiusCli_V32_work
```

Par défaut, le point d'entrée VBA est `Zyva`. Un autre point d'entrée peut être
fourni avec `-EntryPoint`.

La comparaison utilise des tolérances numériques. Elle ignore explicitement
`OutputSynt.BilanNnonOK` et `OutputD_1.DrainageON`, deux indicateurs discrets
sensibles aux valeurs situées exactement au voisinage d'un seuil flottant.

## Intégration continue

Le workflow GitHub Actions est lancé automatiquement à chaque push et pull
request. Il réalise sous Linux x64 :

1. l'installation du SDK .NET 9 ;
2. la restauration des dépendances ;
3. la compilation Release ;
4. la publication du CLI ;
5. le test `CelsiusCli --help`.

La régression complète n'est pas exécutée sur les runners GitHub standards, car
elle nécessite Microsoft Access. Elle doit être lancée localement ou sur un
runner Windows auto-hébergé disposant d'Access et de la base V32.

## Règles de contribution

- Ne pas modifier manuellement `CelsiusCli_V32_work/Converted/*.vb` sans
  répercuter la correction dans la source VBA.
- Ne jamais ajouter les bases Access, les bases SQLite de travail ni les sorties
  de simulation au dépôt.
- Exécuter au minimum la compilation Linux avant un commit.
- Exécuter la régression Access/VB.NET après une modification du moteur
  scientifique ou du convertisseur.
- Examiner ensemble les changements `.bas` et les changements générés `.vb`.

Des détails supplémentaires sont disponibles dans
`CelsiusCli_V32_work/README.md` et `celsius_pipeline_v32/README.md`. La carte
des processus biophysiques et de leurs interactions est disponible dans
[`CELSIUS_V32_Process_Graph.md`](CelsiusCli_V32_work/docs/CELSIUS_V32_Process_Graph.md).

## Licence

Voir le fichier `LICENSE` à la racine du dépôt.
