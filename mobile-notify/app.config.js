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
  if (fs.existsSync(gsPath)) {
    next.android = { ...next.android, googleServicesFile: "./google-services.json" };
  } else if (next.android?.googleServicesFile) {
    const { googleServicesFile, ...androidRest } = next.android;
    next.android = androidRest;
  }
  return next;
};
