import { defineConfig } from "vite";

import { createAppViteConfig, getOfficeHttpsServerOptions } from "./src/devServerConfig";

export default defineConfig(async () => {
  const loadHttpsOptions = process.env.VITEST
    ? async () => ({})
    : getOfficeHttpsServerOptions;

  return createAppViteConfig(loadHttpsOptions);
});
