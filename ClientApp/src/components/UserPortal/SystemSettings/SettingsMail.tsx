// ---------------------------------------------------------------------------
// Helper for sending an email from DMS.
// The email always opens directly in Outlook on the web (new browser tab)
// with the recipient, subject and message already filled in.
// ---------------------------------------------------------------------------

export interface IMailMessage {
  to: string;
  subject: string;
  body: string;
}

// Builds the Outlook on the web "compose" link for a message.
const buildOutlookWebUrl = (message: IMailMessage): string => {
  const to = encodeURIComponent(message.to);
  const subject = encodeURIComponent(message.subject);
  const body = encodeURIComponent(message.body);
  return `https://outlook.office.com/mail/deeplink/compose?to=${to}&subject=${subject}&body=${body}`;
};

// Opens Outlook on the web in a new browser tab. Call this directly from a click handler.
export const openOutlookOnWeb = (message: IMailMessage): void => {
  window.open(buildOutlookWebUrl(message), '_blank', 'noopener,noreferrer');
};
