// api/send-report.js
const allowOrigin = "*"; // remplace par ton domaine quand tout marche

function setCors(res) {
  res.setHeader("Access-Control-Allow-Origin", allowOrigin);
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS, GET");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
}

module.exports = async (req, res) => {
  setCors(res);

  if (req.method === "OPTIONS") return res.status(200).end();

  if (req.method === "GET") {
    return res.status(200).json({
      ok: true,
      hint: "Use POST with JSON body",
      env: {
        M365_TENANT_ID: !!process.env.M365_TENANT_ID,
        M365_CLIENT_ID: !!process.env.M365_CLIENT_ID,
        M365_CLIENT_SECRET: !!process.env.M365_CLIENT_SECRET,
        M365_SENDER_UPN: !!process.env.M365_SENDER_UPN,
      },
      time: new Date().toISOString(),
    });
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Use POST" });
  }

  try {
    let body = req.body;
    if (typeof body === "string") {
      try { body = JSON.parse(body); }
      catch (e) { return res.status(400).json({ error: "Invalid JSON", details: e.message }); }
    }

    const {
      to,
      subject = "Rapport d'audit",
      text = "Bonjour, veuillez trouver le rapport en pièce jointe.",
      filename = "rapport.pdf",
      contentBase64, // PDF en base64 (sans "data:application/pdf;base64,")
    } = body || {};

    if (!to || !contentBase64) {
      return res.status(400).json({ error: "Champs requis: to, contentBase64" });
    }

    const tenantId = process.env.M365_TENANT_ID;
    const clientId = process.env.M365_CLIENT_ID;
    const clientSecret = process.env.M365_CLIENT_SECRET;
    const senderUpn = process.env.M365_SENDER_UPN;

    const missing = [];
    if (!tenantId) missing.push("M365_TENANT_ID");
    if (!clientId) missing.push("M365_CLIENT_ID");
    if (!clientSecret) missing.push("M365_CLIENT_SECRET");
    if (!senderUpn) missing.push("M365_SENDER_UPN");
    if (missing.length) return res.status(500).json({ error: "Env vars missing", missing });

    const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    const tokenBody = new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials",
    });

    const tokenRes = await fetch(tokenUrl, { method: "POST", body: tokenBody });
    const tokenTxt = await tokenRes.text();
    if (!tokenRes.ok) return res.status(tokenRes.status).json({ error: "Token error", details: tokenTxt });
    const { access_token } = JSON.parse(tokenTxt);

    const graphUrl = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(senderUpn)}/sendMail`;
    const payload = {
      message: {
        subject,
        body: { contentType: "Text", content: text },
        toRecipients: [{ emailAddress: { address: to } }],
        attachments: [
          {
            "@odata.type": "#microsoft.graph.fileAttachment",
            name: filename,
            contentType: "application/pdf",
            contentBytes: contentBase64,
          },
        ],
      },
      saveToSentItems: true,
    };

    const mailRes = await fetch(graphUrl, {
      method: "POST",
      headers: { Authorization: `Bearer ${access_token}`, "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const mailTxt = await mailRes.text();
    if (!mailRes.ok) return res.status(mailRes.status).json({ error: "Graph sendMail error", details: mailTxt });

    return res.status(200).json({ success: true });
  } catch (e) {
    console.error("send-report error:", e);
    return res.status(500).json({ error: e?.message || String(e) });
  }
};

