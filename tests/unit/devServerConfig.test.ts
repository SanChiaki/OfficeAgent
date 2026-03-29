// @vitest-environment node

import { describe, expect, it } from "vitest";

import { createAppViteConfig } from "../../src/devServerConfig";

describe("createAppViteConfig", () => {
  it("uses Office dev certificates for the localhost HTTPS server", async () => {
    const httpsOptions = {
      key: "key",
      cert: "cert"
    };

    const config = await createAppViteConfig(async () => httpsOptions);

    expect(config.server?.https).toBe(httpsOptions);
    expect(config.server?.host).toBe("localhost");
    expect(config.server?.port).toBe(3000);
  });
});
