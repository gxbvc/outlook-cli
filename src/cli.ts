import { Command } from "commander";
import { runAuthFlow } from "./auth.js";
import { registerEmails } from "./commands/emails.js";
import { registerDrafts } from "./commands/drafts.js";
import { registerFolders } from "./commands/folders.js";
import { registerMe } from "./commands/me.js";
import { registerCalendar } from "./commands/calendar.js";
import { registerMeetings } from "./commands/meetings.js";

const program = new Command();

program
  .name("outlook-cli")
  .description("CLI for Microsoft Outlook via Graph API")
  .version("1.0.0")
  .option("--user <email>", "Act as this mailbox (shared mailbox / Send As)")
  .hook("preAction", (_thisCommand, actionCommand) => {
    const user = actionCommand.optsWithGlobals().user;
    if (user) process.env.OUTLOOK_USER = user;
  });

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
registerCalendar(program);
registerMeetings(program);

program.parse();
