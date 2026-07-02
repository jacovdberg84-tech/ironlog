import { runPlannerAgent } from "./agents/planner.js";
import { runCoderAgent } from "./agents/coder.js";

export async function runAgent(agentName, context) {
  switch (String(agentName || "").toLowerCase()) {
    case "planner":
      return runPlannerAgent(context);
    case "coder":
      return runCoderAgent(context);
    default:
      return {
        ok: false,
        summary: `Unknown agent: ${agentName}`,
        details: { agent: agentName },
      };
  }
}

export { runPlannerAgent, runCoderAgent };
