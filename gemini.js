// netlify/functions/gemini.js
// Serverless proxy — forwards requests to Gemini API
// Runs on Netlify's servers, so no CORS issues

exports.handler = async (event) => {
  // Only allow POST
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Method Not Allowed" };
  }

  try {
    const { apiKey, prompt, pdfB64, mimeType, textOnly } = JSON.parse(event.body);

    if (!apiKey || !prompt) {
      return { statusCode: 400, body: JSON.stringify({ error: "Missing apiKey or prompt" }) };
    }

    const model = "gemini-2.0-flash";
    const url   = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

    // Build request parts
    const parts = [{ text: prompt }];
    if (!textOnly && pdfB64) {
      parts.push({ inline_data: { mime_type: mimeType || "application/pdf", data: pdfB64 } });
    }

    const geminiResp = await fetch(url, {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body:    JSON.stringify({
        contents: [{ parts }],
        generationConfig: { temperature: 0.1, maxOutputTokens: 8000 },
      }),
    });

    const data = await geminiResp.json();

    if (!geminiResp.ok) {
      return {
        statusCode: geminiResp.status,
        headers: { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" },
        body: JSON.stringify({ error: data.error?.message || "Gemini API error" }),
      };
    }

    const text = data.candidates?.[0]?.content?.parts?.[0]?.text || "";

    return {
      statusCode: 200,
      headers: {
        "Content-Type": "application/json",
        "Access-Control-Allow-Origin":  "*",
        "Access-Control-Allow-Headers": "Content-Type",
      },
      body: JSON.stringify({ text }),
    };

  } catch (err) {
    return {
      statusCode: 500,
      headers: { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" },
      body: JSON.stringify({ error: err.message }),
    };
  }
};
