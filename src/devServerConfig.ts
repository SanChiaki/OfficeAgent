import type { ServerOptions } from "node:https";

import react from "@vitejs/plugin-react";
import * as devCerts from "office-addin-dev-certs";
import type { UserConfig } from "vite";

export async function getOfficeHttpsServerOptions(): Promise<ServerOptions> {
  return devCerts.getHttpsServerOptions();
}

export async function createAppViteConfig(
  loadHttpsOptions: () => Promise<ServerOptions> = getOfficeHttpsServerOptions
): Promise<UserConfig> {
  const httpsOptions = await loadHttpsOptions();

  return {
    plugins: [react()],
    server: {
      host: "localhost",
      https: httpsOptions,
      port: 3000,
      strictPort: true
    },
    build: {
      outDir: "dist",
      sourcemap: true,
      rollupOptions: {
        input: {
          taskpane: "taskpane.html"
        }
      }
    },
    test: {
      environment: "jsdom",
      setupFiles: ["./tests/setup.ts"]
    }
  };
}
