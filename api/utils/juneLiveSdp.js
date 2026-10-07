// IRONLOG/api/utils/juneLiveSdp.js
// OpenAI's WARP/SNAP optimisation is experimental in browsers. A normal
// browser offer uses the standard SCTP data-channel attribute, so a SNAP-only
// answer cannot be applied by browsers without the matching origin trial.

function hasLine(sdp, expression) {
  return expression.test(String(sdp || ""));
}

export function makeJuneLiveAnswerBrowserCompatible({ offer, answer }) {
  const remoteAnswer = String(answer || "");
  // Leave modern/SNAP-capable negotiations alone. The compatibility path is
  // only for a normal browser offer receiving a SNAP-only response.
  if (
    !hasLine(remoteAnswer, /^a=sctp-init:[^\r\n]*/m)
    || hasLine(offer, /^a=sctp-init:[^\r\n]*/m)
  ) return remoteAnswer;

  const newline = remoteAnswer.includes("\r\n") ? "\r\n" : "\n";
  const hasStandardPort = hasLine(remoteAnswer, /^a=sctp-port:\d+/m);
  return remoteAnswer.replace(
    /^a=sctp-init:[^\r\n]*(?:\r?\n|$)/gm,
    hasStandardPort ? "" : `a=sctp-port:5000${newline}`,
  );
}
