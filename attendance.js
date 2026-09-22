const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyyvwWzCmcHUkl-C0BTm_2Q8VtL5B0UEIEkClSOqPGlozvJFeR6d0lx6tNwrhSwq47RCw/exec";
const API_TOKEN = "7fQC2Sx7d6RHKm-73gAE2FZjpPUebFIdcy1YxUg8slo";

async function fetchAppsScript(url, options = {}, maxRedirects = 5) {
  let currentUrl = url;
  const method = options.method || "GET";
  const body = options.body;

  for (let i = 0; i <= maxRedirects; i++) {
    const response = await fetch(currentUrl, {
      method,
      headers: { ...(options.headers || {}) },
      body,
      redirect: "manual",
    });

    if ([301, 302, 303, 307, 308].includes(response.status)) {
      const location = response.headers.get("location");
      if (!location) return response;
      currentUrl = new URL(location, currentUrl).toString();
      continue;
    }

    return response;
  }

  throw new Error("Too many redirects while contacting Apps Script.");
}

async function readResponse(response) {
  const text = await response.text();

  try {
    return JSON.parse(text);
  } catch (_) {
    throw new Error(
      "Apps Script returned an unexpected response. " +
      "Make sure the updated Apps Script deployment is active."
    );
  }
}

module.exports = async function handler(req, res) {
  try {
    if (req.method === "GET") {
      const date = String(req.query.date || "").trim();

      if (!date) {
        return res.status(400).json({
          success: false,
          error: "Date is required."
        });
      }

      const url =
        APPS_SCRIPT_URL +
        "?action=get" +
        "&date=" + encodeURIComponent(date) +
        "&token=" + encodeURIComponent(API_TOKEN);

      const response = await fetchAppsScript(url);
      const result = await readResponse(response);

      if (result && result.error) {
        return res.status(400).json(result);
      }

      return res.status(200).json(result);
    }

    if (req.method === "POST") {
      const payload =
        typeof req.body === "string"
          ? JSON.parse(req.body)
          : (req.body || {});

      if (
        payload.action !== "saveOne" &&
        payload.action !== "saveAll"
      ) {
        return res.status(400).json({
          success: false,
          error: "Invalid API action."
        });
      }

      payload.token = API_TOKEN;

      const response = await fetchAppsScript(
        APPS_SCRIPT_URL,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload)
        }
      );

      const result = await readResponse(response);

      if (result && result.success === false) {
        return res.status(400).json(result);
      }

      return res.status(200).json(result);
    }

    res.setHeader("Allow", "GET, POST");
    return res.status(405).json({
      success: false,
      error: "Method not allowed."
    });

  } catch (error) {
    return res.status(500).json({
      success: false,
      error: error.message || "Server error."
    });
  }
};
