# Edge-SaveAs-Check

Vérifie, depuis Power Automate Desktop (PAD), que le téléchargement en attente d'une page Edge
propose bien la fenêtre Windows « Enregistrer sous » : le script retrouve l'onglet (URL + titre),
ouvre le panneau Téléchargements d'Edge, clique sur « Enregistrer sous » et attend la fenêtre.

## Installer le script dans un flux PAD

**Option A (recommandée) : version autonome, rien à installer sur les postes.**
Le script voyage avec le flux : ceux qui exécutent le flux n'ont rien à copier.

1. Dans le flux, créer les variables `BrowserPageURL` et `BrowserPageTitle` (action « Définir la
   variable »). Jokers acceptés, par exemple `*lims.eu.eurofinsbpt.com/elimsMetrology*` et
   `*eLIMS-BPT*`. Apostrophes autorisées.
2. Ajouter une action « Exécuter un script PowerShell » et y coller tout le contenu de
   [`Edge-SaveAs-Check.PAD-standalone.ps1`](Edge-SaveAs-Check.PAD-standalone.ps1) (bouton « Copy raw file »).
3. Dans cette action, décocher « Échec après expiration du délai » (ou mettre au moins 120 s).

**Option B : script à un emplacement partagé.** Un dossier réseau lisible par tous les
utilisateurs contient `Edge-SaveAs-Check.ps1` ; coller le lanceur
[`Edge-SaveAs-Check.PAD.ps1`](Edge-SaveAs-Check.PAD.ps1) dans l'action et régler `$ScriptPath`.
Une seule copie à mettre à jour, mais le flux dépend de ce partage.

## Résultat

`%PowershellOutput%` contient un code de résultat (suivi d'un retour à la ligne : le rogner) ;
en cas d'échec, `%ScriptError%` contient le message.

| Code | Signification |
|---|---|
| 0 | La fenêtre « Enregistrer sous » s'est ouverte |
| 1 | Erreur inattendue (voir le journal) |
| 2 | Edge n'est pas lancé |
| 3 | Aucune page Edge ne correspond à l'URL et au titre |
| 4 | Bouton Téléchargements introuvable |
| 5 | Bouton « Enregistrer sous » introuvable dans le panneau |
| 6 | La fenêtre Windows ne s'est pas ouverte |
| 8 | PowerShell en mode restreint (stratégie de sécurité) : script inutilisable |
| 9 | Variable PAD vide (`BrowserPageURL` ou `BrowserPageTitle`) |
| 10 à 13 | Lanceur seulement : script introuvable, illisible, invalide ou incomplet |

## Conditions

- Windows PowerShell 5.1 (présent par défaut sous Windows 10 et 11), Microsoft Edge lancé dans la
  **session de l'utilisateur** connecté : le script ne fonctionne pas sous le compte SYSTEM ni en
  session déconnectée.
- Un téléchargement doit être en attente dans l'Edge visé (le script est fait pour le moment où
  l'application web vient de déclencher le téléchargement).
- Aucun droit administrateur.

## Langue du navigateur

Les éléments sont reconnus par leur **classe UI Automation** (`ToolbarDownloadsButton`,
`saveAs`, `OmniboxViewViews`), qui n'est pas traduite, et la fenêtre « Enregistrer sous » par sa
classe de boîte de dialogue Windows (`#32770`) issue d'Edge, sans lire son titre. Les noms
français et anglais ne servent que de repli. Le titre de la page (`BrowserPageTitle`) est celui
de la page, que vous fournissez.

État de la validation : testé sur Edge en français (Windows 11, build 26100). Les autres langues
sont prévues mais pas encore testées, et Edge peut changer ces noms de classe d'une version à
l'autre : en cas d'échec (code 4 ou 5), le journal liste les boutons vus avec leur classe
(« Diagnostic - … nom[classe] »), ce qui permet de corriger `-DownloadsButtonClasses` et
`-SaveAsButtonClasses`.

## Arrière-plan

Le script ne déplace pas le curseur, n'envoie aucune frappe au clavier et ne met pas Edge au
premier plan. Il utilise UI Automation, l'accessibilité MSAA et des messages Windows postés
directement aux fenêtres d'Edge (jamais à une autre application, même si elle recouvre Edge).
Ce que l'utilisateur voit quand même, parce que c'est Edge qui l'affiche : le panneau
Téléchargements qui s'ouvre, l'onglet cible qui devient actif s'il ne l'était pas, et la fenêtre
« Enregistrer sous » (laissée ouverte, ou fermée avec `CloseDialog = $true` : aucun fichier n'est
alors enregistré). Un vrai clic souris n'est utilisé qu'avec `AllowRealMouse` (désactivé).

## Journaux

Un fichier par exécution dans `%LOCALAPPDATA%\EdgeSaveAsCheck\Logs` (profil de chaque
utilisateur), nommé avec l'utilisateur et l'heure ; les journaux de plus de 30 jours sont
supprimés automatiquement. Il indique chaque étape, la méthode de clic qui a fonctionné et se
termine toujours par une ligne `RESULTAT`.

## Mettre à jour

Modifier `Edge-SaveAs-Check.ps1`, puis relancer `Build-PadStandalone.ps1` pour régénérer la
version autonome (il vérifie la syntaxe, l'absence de signe pourcent parasite et de BOM).
Le fichier source ne doit contenir aucun signe pourcent : PAD l'interpréterait comme une variable.
