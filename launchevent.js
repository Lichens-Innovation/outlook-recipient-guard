/*
 * Recipient Guard — handler OnMessageSend (Smart Alerts)
 *
 * Lit la configuration dans les roamingSettings de la boîte (éditée via le volet
 * « Recipient Guard » du ruban), compare les destinataires To/Cc/Bcc et, en cas de
 * correspondance, bloque l'envoi avec une alerte « Envoyer quand même / Ne pas envoyer ».
 *
 * Code volontairement sans async/await ni optional chaining : il doit tourner dans le
 * runtime JavaScript-only d'Outlook classique pour Windows.
 */

var RG_SETTINGS_KEY = "rg.config";

// Valeurs par défaut si rien n'a encore été enregistré via le volet.
var RG_DEFAULT_CONFIG = {
  rules: "",              // une règle par ligne (voir parseRules)
  warnExternal: false,    // alerter pour tout domaine hors internalDomains
  internalDomains: ""     // ex. "lichens.ai, lichens.com"
};

function rgLoadConfig() {
  var cfg = null;
  try {
    cfg = Office.context.roamingSettings.get(RG_SETTINGS_KEY);
  } catch (e) {
    cfg = null;
  }
  var out = {};
  for (var k in RG_DEFAULT_CONFIG) {
    out[k] = cfg && cfg[k] !== undefined ? cfg[k] : RG_DEFAULT_CONFIG[k];
  }
  return out;
}

/*
 * Syntaxe des règles (insensible à la casse) :
 *   jean.tremblay@gmail.com      adresse exacte
 *   @ancien-client.com           tout le domaine (équivaut à *@ancien-client.com)
 *   jean*@*                      jokers * acceptés partout
 *   # commentaire                ignoré
 *   x@y.com | Adresse perso      texte après « | » = raison affichée dans l'alerte
 */
function rgParseRules(text) {
  var rules = [];
  (text || "").split(/\r?\n/).forEach(function (raw) {
    var line = raw.trim();
    if (!line || line.charAt(0) === "#") return;
    var reason = "";
    var pipe = line.indexOf("|");
    if (pipe >= 0) {
      reason = line.slice(pipe + 1).trim();
      line = line.slice(0, pipe).trim();
    }
    if (!line) return;
    var pattern = line.toLowerCase();
    if (pattern.charAt(0) === "@") pattern = "*" + pattern;
    var regex = new RegExp(
      "^" + pattern.replace(/[.+?^${}()[\]\\]/g, "\\$&").replace(/\*/g, ".*") + "$"
    );
    rules.push({ pattern: line, reason: reason, regex: regex });
  });
  return rules;
}

function rgParseDomains(text) {
  return (text || "")
    .split(/[\s,;]+/)
    .map(function (d) { return d.trim().toLowerCase().replace(/^@/, ""); })
    .filter(function (d) { return d.length > 0; });
}

function rgGetRecipients(field) {
  return new Promise(function (resolve) {
    if (!field || !field.getAsync) return resolve([]);
    field.getAsync(function (res) {
      resolve(res.status === Office.AsyncResultStatus.Succeeded ? res.value || [] : []);
    });
  });
}

function rgFindMatches(recipients, cfg) {
  var rules = rgParseRules(cfg.rules);
  var internal = rgParseDomains(cfg.internalDomains);
  var hits = [];
  var seen = {};

  recipients.forEach(function (r) {
    var addr = (r.emailAddress || "").trim().toLowerCase();
    if (!addr || seen[addr]) return;

    for (var i = 0; i < rules.length; i++) {
      if (rules[i].regex.test(addr)) {
        seen[addr] = true;
        hits.push({ address: addr, why: rules[i].reason || rules[i].pattern });
        return;
      }
    }

    if (cfg.warnExternal && internal.length) {
      var domain = addr.split("@")[1] || "";
      var isInternal = internal.some(function (d) {
        return domain === d || domain.slice(-(d.length + 1)) === "." + d;
      });
      if (!isInternal) {
        seen[addr] = true;
        hits.push({ address: addr, why: "externe" });
      }
    }
  });
  return hits;
}

function rgBuildMessage(hits) {
  var lines = hits.map(function (h) { return h.address + " (" + h.why + ")"; });
  var msg = "Destinataire(s) à vérifier : " + lines.join(", ") + ". Envoyer quand même ?";
  // Limite Outlook : 500 caractères.
  if (msg.length > 500) {
    msg = "Destinataire(s) à vérifier (" + hits.length + ") : " + lines.join(", ");
    msg = msg.slice(0, 470) + "… Envoyer quand même ?";
  }
  return msg;
}

function onMessageSendHandler(event) {
  try {
    var item = Office.context.mailbox.item;
    var cfg = rgLoadConfig();

    Promise.all([
      rgGetRecipients(item.to),
      rgGetRecipients(item.cc),
      rgGetRecipients(item.bcc)
    ])
      .then(function (lists) {
        var all = lists[0].concat(lists[1], lists[2]);
        var hits = rgFindMatches(all, cfg);
        if (hits.length === 0) {
          event.completed({ allowEvent: true });
        } else {
          event.completed({ allowEvent: false, errorMessage: rgBuildMessage(hits) });
        }
      })
      .catch(function () {
        // En cas d'erreur, ne jamais bloquer l'utilisateur.
        event.completed({ allowEvent: true });
      });
  } catch (e) {
    event.completed({ allowEvent: true });
  }
}

// Exposé pour le volet (bouton « Tester »).
if (typeof window !== "undefined") {
  window.rgParseRules = rgParseRules;
  window.rgFindMatches = rgFindMatches;
  window.RG_SETTINGS_KEY = RG_SETTINGS_KEY;
  window.rgLoadConfig = rgLoadConfig;
}

// Enregistrement du handler (requis sur toutes les plateformes).
try {
  if (typeof Office !== "undefined" && Office.actions) {
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  }
} catch (e) {
  // Ignoré dans le volet (taskpane), où Office.actions n'est pas pertinent.
}
