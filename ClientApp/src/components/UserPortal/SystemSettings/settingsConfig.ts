// ---------------------------------------------------------------------------
// Central content configuration for the Settings modules.
// Edit the text, version history and contact details here; the screens
// (About, Help & Support, Privacy Policy, Version & Release Information)
// render everything from this file, so no screen needs to be touched.
// ---------------------------------------------------------------------------

export const APP_NAME = 'DMS';
export const APP_FULL_NAME = 'Document Management System';
export const APP_TAGLINE = 'Secure document control, approvals and process management on Microsoft 365.';

export const SUPPORT_EMAIL = 'support@speteq.com';
export const SUPPORT_PHONE = '+91 95866-61642';

// Choices shown in the "Report an issue" form on the Help & Support page.
export const ISSUE_TYPES: string[] = [
  'Something is not working',
  'Page or data is not loading',
  'Approval problem',
  'Document problem (upload, download, share)',
  'Access or permission problem',
  'Wrong or missing information',
  'Suggestion or improvement'
];

export const AFFECTED_AREAS: string[] = [
  'Dashboard',
  'Documents',
  'Approvals',
  'Processes',
  'Settings',
  'Other'
];

export const ISSUE_PRIORITIES: string[] = [
  'Low - can wait',
  'Medium - slows my work',
  'High - I cannot continue my work'
];

// ---- About ----------------------------------------------------------------

export const ABOUT_DESCRIPTION =
  'DMS is a document management application built on SharePoint Framework. ' +
  'It keeps your organisation\'s documents, approval workflows and business processes in one place, ' +
  'using your existing Microsoft 365 sign-in and permissions.';

export interface IAboutHighlight {
  iconName: string;
  title: string;
  text: string;
}

export const ABOUT_HIGHLIGHTS: IAboutHighlight[] = [
  {
    iconName: 'file',
    title: 'Document control',
    text: 'Store, organise and find documents in SharePoint libraries with clear status tracking.'
  },
  {
    iconName: 'check',
    title: 'Multi-level approvals',
    text: 'Route documents through L1, L2 and L3 approvers with a visible approval stage.'
  },
  {
    iconName: 'processes',
    title: 'Process management',
    text: 'Define business processes and keep every related document under its process code.'
  },
  {
    iconName: 'shield',
    title: 'Secure by design',
    text: 'Access follows Microsoft 365 identity and SharePoint permissions. Nothing is stored outside your tenant.'
  }
];

// ---- Help & Support ---------------------------------------------------------

export interface IFaqItem {
  question: string;
  answer: string;
}

export const FAQ_ITEMS: IFaqItem[] = [
  {
    question: 'How do I send a document for approval?',
    answer: 'Open the Documents page, select the document and use the send-for-approval action. The document then appears on the first approver\'s Approvals page.'
  },
  {
    question: 'Why can I not see the Settings menu?',
    answer: 'Settings is available only to DMS administrators. Contact IT support if you need access.'
  },
  {
    question: 'My profile photo or details look wrong. How do I change them?',
    answer: 'Profile details come directly from your Microsoft 365 account and cannot be edited in DMS. Update them from your Microsoft account page and refresh DMS.'
  },
  {
    question: 'A document is stuck in an approval stage. What should I do?',
    answer: 'Check the Approvals page to see which approver currently holds the document, then contact that approver. If they are unavailable, ask IT support to review the workflow configuration.'
  }
];

// ---- Privacy Policy -----------------------------------------------------------

export const PRIVACY_LAST_UPDATED = '2026-10-21';

export interface IPrivacySection {
  iconName: string;
  title: string;
  points: string[];
}

export const PRIVACY_SECTIONS: IPrivacySection[] = [
  {
    iconName: 'user',
    title: 'Information we access',
    points: [
      'Your Microsoft 365 profile: name, email address, job title, department and profile photo.',
      'Documents, folders and list items in the SharePoint site where DMS is installed.',
      'Approval and process records created through DMS.'
    ]
  },
  {
    iconName: 'file',
    title: 'How the information is used',
    points: [
      'To show you who you are signed in as and to personalise the application.',
      'To display, route and track documents and approvals that you are allowed to see.',
      'To send notifications related to documents you own or must approve.'
    ]
  },
  {
    iconName: 'shield',
    title: 'Access and storage',
    points: [
      'All data stays inside your Microsoft 365 tenant (SharePoint and Microsoft Graph).',
      'DMS never grants you more access than your SharePoint permissions already allow.',
      'Display preferences are stored only in your own browser.'
    ]
  },
  {
    iconName: 'mail',
    title: 'Your choices and contact',
    points: [
      'Profile information is managed in your Microsoft account, not in DMS.',
      'You can clear saved display preferences at any time from Appearance & Preferences.',
      'For privacy questions, contact your IT administrator or the support address on the Help & Support page.'
    ]
  }
];

// ---- Version & Release Information ---------------------------------------------

export type ReleaseType = 'Major' | 'Minor' | 'Patch';

export interface IRelease {
  version: string;
  date: string; // ISO format, YYYY-MM-DD
  type: ReleaseType;
  highlights: string[];
}

// Newest release first. The first entry is treated as the current version.
export const RELEASES: IRelease[] = [
  {
    version: '1.0.0',
    date: '2026-10-1',
    type: 'Major',
    highlights: [
      'Dashboard with document and approval summaries.',
      'Document library with status tracking and sharing.',
      'Multi-level approval workflow (L1, L2, L3).',
      'Process management with process-wise folders.',
      'System Settings with profile, preferences, help and release information.'
    ]
  }
];
