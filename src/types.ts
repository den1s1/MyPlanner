export type ProjectStatus = 'active' | 'paused' | 'draft'
export type Priority = 'high' | 'medium' | 'low'
export type TaskStatus = 'planned' | 'in_progress' | 'blocked' | 'done'
export type TaskSource = 'manual' | 'meeting' | 'outlook' | 'youtrack'

export interface Project {
  id: string
  code: string
  name: string
  client: string
  status: ProjectStatus
  color: string
  deadline: string
  progress: number
  contractNumber: string
  contractName: string
  correspondencePath: string
}

export interface Task {
  id: string
  title: string
  projectId: string
  due: string
  priority: Priority
  done: boolean
  description: string
  status: TaskStatus
  assignee: string
  source: TaskSource
  contractStageId: string
  meetingId: string
}

export interface Note {
  id: string
  text: string
  projectId?: string
  createdAt: string
}

export interface CapturedEmail {
  id: string
  outlookItemId: string
  subject: string
  senderName: string
  senderEmail: string
  recipients: string
  receivedAt: string
  excerpt: string
  attachments: CapturedAttachment[]
  capturedAt: string
  processed: boolean
}

export interface CapturedAttachment {
  id: string
  name: string
  path: string
  size: number
}

export type TrackingStatus = 'waiting' | 'reminded' | 'completed' | 'cancelled'

export interface Tracking {
  id: string
  projectId: string
  subject: string
  recipient: string
  sentDate: string
  checkDate: string
  status: TrackingStatus
  comment: string
  sourceEmailId: string
}

export interface QuickLink {
  id: string
  label: string
  target: string
  projectId?: string
  kind: 'web' | 'file'
}

export type ContractStageStatus = 'not_started' | 'started' | 'suspended' | 'work_completed' | 'act_sent' | 'act_signed' | 'closed'
export type PaymentStatus = 'paid' | 'unpaid'

export interface ContractStage {
  id: string
  projectId: string
  number: string
  name: string
  scope: string
  startDate: string
  endDate: string
  status: ContractStageStatus
  suspensionLetter: string
  cost: number
  advancePercent: number
  advanceStatus: PaymentStatus
  finalPaymentStatus: PaymentStatus
  comment: string
}

export interface MeetingAction {
  id: string
  title: string
  assignee: string
  due: string
  taskId: string
}

export interface Meeting {
  id: string
  projectId: string
  date: string
  time: string
  format: 'online' | 'offline' | 'hybrid'
  topic: string
  participants: string
  agenda: string
  protocol: string
  decisions: string
  actions: MeetingAction[]
}

export interface PlannerData {
  projects: Project[]
  tasks: Task[]
  notes: Note[]
  links: QuickLink[]
  contractStages: ContractStage[]
  meetings: Meeting[]
  inbox: CapturedEmail[]
  trackings: Tracking[]
  integrations: {
    worklogExcelPath: string
    youtrackTokenPath: string
    outlookBridgeKey: string
  }
}
