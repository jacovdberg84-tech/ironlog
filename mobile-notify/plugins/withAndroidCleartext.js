const { withAndroidManifest, withDangerousMod } = require("@expo/config-plugins");
const fs = require("fs");
const path = require("path");

module.exports = function withAndroidCleartext(config) {
  config = withDangerousMod(config, [
    "android",
    async (cfg) => {
      const root = cfg.modRequest.platformProjectRoot;
      const xmlDir = path.join(root, "app", "src", "main", "res", "xml");
      fs.mkdirSync(xmlDir, { recursive: true });
      const file = path.join(xmlDir, "network_security_config.xml");
      const body = `<?xml version="1.0" encoding="utf-8"?>
<network-security-config>
  <base-config cleartextTrafficPermitted="true">
    <trust-anchors>
      <certificates src="system" />
    </trust-anchors>
  </base-config>
</network-security-config>
`;
      fs.writeFileSync(file, body);
      return cfg;
    },
  ]);

  return withAndroidManifest(config, async (cfg) => {
    const application = cfg.modResults.manifest.application?.[0];
    if (application?.$) {
      application.$["android:networkSecurityConfig"] = "@xml/network_security_config";
      application.$["android:usesCleartextTraffic"] = "true";
    }
    return cfg;
  });
};
