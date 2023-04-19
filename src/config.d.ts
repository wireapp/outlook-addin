interface Config {
  addInBaseUrl: string;
  apiBaseUrl: string;
  authorizeUrl: string;
  clientId: string;
}

declare const config: Config;

export default config;
