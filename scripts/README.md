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
4. **Premier essai** : exécuter le flux une fois sur un poste avec Edge ouvert sur la page cible.
   Le fichier fait environ 54 Ko (code seul, sans commentaires, ASCII pur) ; si PAD le refusait
   ou le tronquait, le lanceur (option B) reste disponible.

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

Les éléments sont reconnus par leur **classe UI Automation**, qui n'est pas traduite :
`ToolbarDownloadsButton` (bouton Téléchargements) et `saveAs` (bouton « Enregistrer sous »).
La fenêtre « Enregistrer sous » est reconnue par sa classe de boîte de dialogue Windows
(`#32770`) issue d'Edge, sans lire son titre. Les noms français et anglais ne servent que de
repli. Le titre de la page (`BrowserPageTitle`) est celui de la page, que vous fournissez ; les
caractères invisibles (espace de largeur nulle, marques de sens d'écriture, par exemple dans le
titre de la fenêtre Edge) sont ignorés dans la comparaison.

État de la validation : testé sur Edge en français (Windows 11, build 26100) pour le bouton
Téléchargements, le bouton « Enregistrer sous » et la fenêtre. La barre d'adresse a
été lue sur cet Edge, mais on ne sait pas si c'est par sa classe (`OmniboxViewViews`) ou par son
nom (repli) : la ligne « Barre d adresse identifiee par » du journal le dit désormais, et si ni
l'un ni l'autre ne correspond, le script prend le premier champ de saisie de la fenêtre. Les
autres langues d'Edge sont prévues mais pas encore testées, et Edge peut changer ces noms de
classe d'une version à l'autre : en cas d'échec (code 4 ou 5), le journal liste les boutons vus
avec leur classe (« Diagnostic - … nom[classe] »), ce qui permet de corriger
`-DownloadsButtonClasses` et `-SaveAsButtonClasses`.

## Arrière-plan

Le script est conçu pour ne pas gêner l'utilisateur : il ne déplace pas le curseur, n'envoie
aucune vraie frappe de clavier et ne demande pas le premier plan. Il utilise UI Automation,
l'accessibilité MSAA et des messages Windows postés directement aux fenêtres d'Edge (jamais à
une autre application, même si elle recouvre Edge). La touche Entrée, dernier essai, est un
message posté à Edge, pas une frappe réelle.

Ce que l'utilisateur voit quand même, parce que c'est Edge qui l'affiche : le panneau
Téléchargements qui s'ouvre, l'onglet cible qui devient actif s'il ne l'était pas, et la fenêtre
« Enregistrer sous » (laissée ouverte, ou fermée avec `CloseDialog = $true` : aucun fichier n'est
alors enregistré), qui peut prendre le premier plan. Un vrai clic souris n'est utilisé qu'avec
`AllowRealMouse` (désactivé).

Ce n'est pas un mode « fenêtre invisible » : le journal note le premier plan et la position du
curseur avant et après le clic (lignes « Etat avant/apres le clic »), pour le vérifier sur chaque
poste. Cas limites : si une autre fenêtre recouvre le panneau d'Edge, le clic peut ne pas
l'atteindre (le script refuse alors de poster à l'autre application et le journal le dit) ; si
Edge tourne en administrateur et pas le script, Windows refuse l'accès ou les messages (codes 3 à 6).

## Journaux

Un fichier par exécution dans `%LOCALAPPDATA%\EdgeSaveAsCheck\Logs` (profil de chaque
utilisateur), nommé avec l'utilisateur, la date, l'heure et l'identifiant du processus ; les
journaux de plus de 30 jours sont supprimés automatiquement. Il indique chaque étape, la méthode
de clic qui a fonctionné et se termine toujours par une ligne `RESULTAT`. Sur un poste partagé
(session distante), seules les fenêtres Edge de la session de l'utilisateur sont considérées.

## Limites connues

- Les variables PAD sont insérées telles quelles dans le script : une valeur contenant un saut de
  ligne, ou une ligne commençant par `'@`, casse l'analyse du script (code 12 du lanceur, erreur
  PowerShell avec la version autonome). Une URL ou un titre de page n'en contient pas.
- Si plusieurs onglets correspondent à l'URL et au titre, le premier trouvé est utilisé (l'onglet
  actif d'une fenêtre avant ses autres onglets). Pour lire l'URL d'un onglet inactif dont le
  titre correspond, le script l'active brièvement, puis revient à l'onglet d'origine.
- Les noms de classe (`ToolbarDownloadsButton`, `saveAs`) sont ceux d'Edge à ce jour.

## Mettre à jour

Modifier `Edge-SaveAs-Check.ps1`, puis relancer `Build-PadStandalone.ps1` pour régénérer la
version autonome. Il retire les commentaires et vérifie que le code est strictement identique
(mêmes jetons), la syntaxe, l'absence de signe pourcent parasite, de BOM et de caractère non ASCII.
Le fichier source ne doit contenir aucun signe pourcent : PAD l'interpréterait comme une variable.

Passage de l'ancienne version (variables `UrlCible` / `TitreCible`, journaux dans `Scripts\Logs`) :
renommer les variables du flux en `BrowserPageURL` et `BrowserPageTitle`, recoller le contenu de la
version autonome ; les journaux sont désormais dans `%LOCALAPPDATA%\EdgeSaveAsCheck\Logs`.
