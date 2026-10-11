// Every email this CLI sends (send, reply, forward) ends with this footer, so
// readers can tell agent-written email from email Christian typed himself. Same
// idea as the robot emoji in slack-cli and imessages-cli. Drafts get no footer,
// because a person reviews and sends those. There is deliberately no opt-out.
export const AGENT_FOOTER = "Sent from my agent \u{1F916}";
const AGENT_FOOTER_HTML = `<div style="color:#888888;font-size:12px">${AGENT_FOOTER}</div>`;

export function withFooterText(text?: string): string {
  if (text?.includes(AGENT_FOOTER)) return text;
  if (!text || !text.trim()) return AGENT_FOOTER;
  return `${text.trimEnd()}\n\n${AGENT_FOOTER}`;
}

export function withFooterHtml(html: string): string {
  if (html.includes(AGENT_FOOTER)) return html;
  const footer = `<br><br>${AGENT_FOOTER_HTML}`;
  return /<\/body>/i.test(html) ? html.replace(/<\/body>/i, `${footer}</body>`) : `${html}${footer}`;
}

// Body content for --html / --text options: HTML wins, as elsewhere in the CLI.
export function footedContent(opts: { html?: string; text?: string }): string {
  return opts.html ? withFooterHtml(opts.html) : withFooterText(opts.text);
}

// Reply and forward "comment" fields. Graph reads the comment as HTML, so
// plain text newlines collapse into one line. Escape the text and use <br>.
export function footedComment(opts: { html?: string; text?: string }): string {
  if (opts.html) return withFooterHtml(opts.html);
  const text = (opts.text || "").trimEnd();
  if (text.includes(AGENT_FOOTER)) return textToHtml(text);
  return text ? withFooterHtml(textToHtml(text)) : AGENT_FOOTER_HTML;
}

function textToHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\n/g, "<br>");
}
