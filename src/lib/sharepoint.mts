import { randomUUID } from "node:crypto";
import { getConfig } from "./config.mts";

function buildMultipartBody(boundary: string) {
  const { sharePointToken } = getConfig();

  return [
    `--${boundary}`,
    `Authorization: Bearer ${sharePointToken}`,
    `X-HTTP-Method-Override: GET`,
    `_post: 1`,
    "",
    `--${boundary}--`,
  ].join("\r\n");
}

/**
 * Authenticated GET via SharePoint's multipart POST convention.
 * Retries on transient errors (429, 500, 502, 503, 504) up to maxRetries times.
 */
export async function spGet(url: string, maxRetries = 2): Promise<Response> {
  const TRANSIENT = new Set([429, 500, 502, 503, 504]);

  for (let attempt = 0; ; attempt++) {
    const boundary = randomUUID();
    let res: Response;
    try {
      res = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": `multipart/form-data;boundary=${boundary}`,
          Origin: "https://loop.cloud.microsoft",
          Referer: "https://loop.cloud.microsoft/",
        },
        body: buildMultipartBody(boundary),
      });
    } catch (err: unknown) {
      if (attempt < maxRetries) {
        const wait = 1000 * 2 ** attempt;
        console.warn(`  [RETRY] Network error (attempt ${attempt + 1}/${maxRetries + 1}), waiting ${wait}ms...`);
        await new Promise((r) => setTimeout(r, wait));
        continue;
      }
      const msg = err instanceof Error ? err.message : String(err);
      throw new Error(`Network error fetching ${url}: ${msg}`, { cause: err });
    }

    if (res.status === 401 || res.status === 403) {
      throw new Error(
        `SharePoint returned ${res.status} — your SHAREPOINT_BEARER_TOKEN has likely expired. ` +
        `Grab a fresh token and update your .env file.`,
      );
    }

    if (TRANSIENT.has(res.status) && attempt < maxRetries) {
      void res.body?.cancel();
      const retryAfter = res.headers.get("Retry-After");
      let wait = 1000 * 2 ** attempt;
      if (retryAfter) {
        const seconds = Number(retryAfter);
        if (Number.isFinite(seconds)) {
          wait = seconds * 1000;
        } else {
          const date = Date.parse(retryAfter);
          if (!Number.isNaN(date)) wait = Math.max(0, date - Date.now());
        }
      }
      console.warn(`  [RETRY] HTTP ${res.status} (attempt ${attempt + 1}/${maxRetries + 1}), waiting ${wait}ms...`);
      await new Promise((r) => setTimeout(r, wait));
      continue;
    }

    return res;
  }
}
