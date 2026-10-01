export interface MicrosoftMailMessage {
  id: string
  source: 'microsoft'
  mailbox: 'inbox' | 'sent' | 'archive'
  changeKey: string | null
  conversationId: string | null
  fromAddress: string | null
  toRecipientCount: number | null
  readObserved: boolean
  hasNonInlineAttachments: boolean | null
  subjectObserved: boolean
  hasSubject: boolean | null
  from: string
  to: string
  cc?: string
  bcc?: string
  subject: string
  preview: string
  body: string
  bodyContentType: 'html' | 'text' | null
  receivedAt: string
  read: boolean
  starred: boolean
  attachments: string[]
}

export interface MicrosoftAccount {
  username: string
  homeAccountId: string
}

export class MicrosoftMailSource {
  constructor(clientId: string, tenant?: string, redirectUri?: string, authClient?: unknown)
  initialize(): Promise<MicrosoftAccount | null>
  connect(options?: { recoverInterrupted?: boolean }): Promise<MicrosoftAccount>
  disconnect(): Promise<void>
  getAccount(): MicrosoftAccount | null
  list(): Promise<MicrosoftMailMessage[]>
  detail(message: MicrosoftMailMessage): Promise<MicrosoftMailMessage>
  archive(message: MicrosoftMailMessage): Promise<void>
  markRead(message: MicrosoftMailMessage, isRead: boolean): Promise<void>
  reply(message: MicrosoftMailMessage, text: string): Promise<void>
}
