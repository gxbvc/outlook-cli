import { Command } from "commander";
import { runAuthFlow } from "./auth.js";
import { registerEmails } from "./commands/emails.js";
import { registerDrafts } from "./commands/drafts.js";
import { registerFolders } from "./commands/folders.js";
import { registerMe } from "./commands/me.js";

const program = new Command();

program
  .name("outlook-cli")
  .description("CLI for Microsoft Outlook via Graph API")
  .version("1.0.0");

program
  .command("auth")
  .description("Authenticate with Microsoft (opens browser)")
  .action(async () => {
    try {
      await runAuthFlow();
    } catch (e: any) {
      console.log(JSON.stringify({ ok: false, error: e.message }));
      process.exit(1);
    }
  });

registerEmails(program);
registerDrafts(program);
registerFolders(program);
registerMe(program);

program.parse();
