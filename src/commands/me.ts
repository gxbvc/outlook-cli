import { Command } from "commander";
import { graphRequest, out, err } from "../graph.js";

export function registerMe(program: Command) {
  program
    .command("me")
    .description("Get current user profile")
    .action(async () => {
      try {
        const data = await graphRequest("/me");
        out(data);
      } catch (e: any) { err(e.message, "ME_FAILED"); }
    });
}
