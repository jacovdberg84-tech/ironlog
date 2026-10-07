import assert from "node:assert/strict";
import test from "node:test";
import { makeJuneLiveAnswerBrowserCompatible } from "../utils/juneLiveSdp.js";

const standardOffer = [
  "v=0",
  "m=application 9 UDP/DTLS/SCTP webrtc-datachannel",
  "a=sctp-port:5000",
  "",
].join("\r\n");

test("replaces a SNAP-only answer line for a standard browser data channel", () => {
  const answer = [
    "v=0",
    "m=application 9 UDP/DTLS/SCTP webrtc-datachannel",
    "a=sctp-init:AQAAHGVPFd8AEAAA/////9tAK2qACAAIgsBAwg==",
    "",
  ].join("\r\n");
  const compatible = makeJuneLiveAnswerBrowserCompatible({ offer: standardOffer, answer });
  assert.match(compatible, /a=sctp-port:5000/);
  assert.doesNotMatch(compatible, /a=sctp-init:/);
});

test("preserves a normal answer and a SNAP-capable negotiation", () => {
  const normal = "v=0\r\na=sctp-port:5000\r\n";
  assert.equal(makeJuneLiveAnswerBrowserCompatible({ offer: standardOffer, answer: normal }), normal);
  const snapOffer = "v=0\r\na=sctp-init:browser-support\r\n";
  const snapAnswer = "v=0\r\na=sctp-init:server-response\r\n";
  assert.equal(makeJuneLiveAnswerBrowserCompatible({ offer: snapOffer, answer: snapAnswer }), snapAnswer);
});
