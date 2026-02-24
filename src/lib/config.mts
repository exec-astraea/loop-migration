type Config = {
    loopToken: string;
    sharePointToken: string;
}

let _config: Config | null = null;

/**
 * Returns and validates the config from env variables
 */
export function getConfig(): Config {
    if (_config) return _config;
    _config = {
        loopToken: requireToken("LOOP_BEARER_TOKEN"),
        sharePointToken: requireToken("SHAREPOINT_BEARER_TOKEN"),
    };
    return _config;
}

function requireToken(name: string): string {
  const raw = process.env[name] ?? "";
  const token = raw.replace(/^Bearer\s+/i, "");
  if (!token) {
    throw new Error(`${name} is missing or empty. See README for setup.`);
  }
  return token;
}
