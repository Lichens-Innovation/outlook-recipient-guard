# Recipient Guard — add-in Outlook

Intercepte l'envoi (Smart Alerts, événement `OnMessageSend`) et affiche une alerte
**« Envoyer quand même / Ne pas envoyer »** si un destinataire (À, Cc, Cci) correspond
à une règle. Protège contre les erreurs d'autocomplétion.

## Fichiers

| Fichier | Rôle |
|---|---|
| `manifest.xml` | À charger dans Outlook via *Add a custom add-in → Add from File* |
| `launchevent.js` | Logique de vérification (handler `onMessageSendHandler`) |
| `commands.html` | Runtime HTML (Outlook web, nouveau Outlook, Mac) |
| `taskpane.html` | Volet de configuration (bouton « Destinataires surveillés » en composition) |
| `assets/` | Icônes |

La liste des règles **n'est pas dans le code** : elle est enregistrée dans les
`roamingSettings` de ta boîte aux lettres (privée, suit ton compte sur tous les clients).

## Installation

1. **Héberger les fichiers en HTTPS.** Le manifeste ne contient que des URL ; Outlook
   charge le JS depuis le web. Le plus simple : un repo GitHub `outlook-recipient-guard`
   avec GitHub Pages activé (Settings → Pages → branche `main`, dossier `/`).
   L'URL attendue par défaut est `https://lichens-innovation.github.io/outlook-recipient-guard/`.
   Autre hôte ? `./set-base-url.sh https://mon-hote/chemin` met à jour le manifeste.
2. Vérifier que `https://…/commands.html` et `https://…/launchevent.js` répondent.
3. Outlook → *Add-ins* → *My add-ins* → **Add a custom add-in → Add from File** →
   `manifest.xml`.
4. Ouvrir un nouveau courriel → ruban → **Destinataires surveillés** → saisir les règles →
   *Enregistrer*. Le bouton *Tester ce brouillon* vérifie les destinataires actuels.

## Syntaxe des règles

```
# commentaire
jean.tremblay@gmail.com | Adresse perso de Jean     ← exacte + raison affichée
@ancien-client.com                                  ← tout le domaine
andre*@gmail.com                                    ← jokers *
```

Option : *Alerter aussi pour tout destinataire externe* + liste de domaines internes
(ex. `lichens.ai`) — les sous-domaines sont considérés internes.

## Comportement

- `SendMode="PromptUser"` : l'alerte propose **Ne pas envoyer** (retour au brouillon)
  ou **Envoyer quand même**. Pour bloquer sans possibilité de contourner, remplacer par
  `SoftBlock` dans le manifeste.
- En cas d'erreur du script, l'envoi n'est jamais bloqué.
- Requiert Mailbox 1.12 : Outlook web, nouveau Outlook (Windows/Mac), Outlook classique
  Windows récent. Pas de support sur mobile (l'envoi passe sans vérification).

## Mise à jour

- Modifier le JS/HTML → pousser sur l'hôte (cache Outlook : vider ou attendre quelques minutes).
- Modifier le manifeste → incrémenter `<Version>`, retirer puis rajouter l'add-in.
