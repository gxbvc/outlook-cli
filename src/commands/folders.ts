import { Command } from "commander";
import { graphRequest, out, err } from "../graph.js";

export function registerFolders(program: Command) {
  const folders = program.command("folders").description("Manage mail folders");

  folders
    .command("list")
    .description("List mail folders")
    .action(async () => {
      try {
        const data = await graphRequest("/me/mailFolders?$top=100");
        out(data.value);
      } catch (e: any) { err(e.message, "LIST_FAILED"); }
    });

  folders
    .command("create")
    .description("Create a mail folder")
    .requiredOption("--name <name>", "Folder name")
    .option("--parent <id>", "Parent folder ID")
    .action(async (opts) => {
      try {
        const base = opts.parent ? `/me/mailFolders/${opts.parent}/childFolders` : "/me/mailFolders";
        const data = await graphRequest(base, {
          method: "POST",
          body: JSON.stringify({ displayName: opts.name }),
        });
        out(data);
      } catch (e: any) { err(e.message, "CREATE_FAILED"); }
    });
}
