/**
 * Preview APK uses a different package so it can install beside production.
 */
const fs = require("fs");
const path = require("path");

module.exports = ({ config }) => {
  let next = { ...config };
  const profile = process.env.EAS_BUILD_PROFILE || "";
  if (profile === "preview") {
    next = {
      ...next,
      name: "IRONLOG Notify (Preview)",
      android: {
        ...next.android,
        package: "com.aml.ironlog.notify.preview",
      },
    };
  }
  const gsPath = path.join(__dirname, "google-services.json");
  const googleServicesFile =
    process.env.GOOGLE_SERVICES_JSON ||
    (fs.existsSync(gsPath) ? "./google-services.json" : null);
  if (googleServicesFile) {
    next.android = { ...next.android, googleServicesFile };
  } else if (next.android?.googleServicesFile) {
    const { googleServicesFile, ...androidRest } = next.android;
    next.android = androidRest;
  }
  return next;
};
