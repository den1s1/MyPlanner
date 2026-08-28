import type { PlannerData } from './types'

export const initialData: PlannerData = {
  projects: [
    { id: 'p1', code: 'PR-024', name: 'Модернизация лаборатории', client: 'АО «Вектор»', status: 'active', color: '#eb7a55', deadline: '2026-10-30', progress: 64, contractNumber: 'Номер договора', contractName: 'Название договора', correspondencePath: '' },
    { id: 'p2', code: 'PR-031', name: 'Проектирование комплекса', client: 'ООО «Горизонт»', status: 'active', color: '#577b73', deadline: '2026-12-18', progress: 28, contractNumber: 'Номер договора', contractName: 'Название договора', correspondencePath: '' },
    { id: 'p3', code: 'PR-036', name: 'Техническое обследование', client: 'ГК «Север»', status: 'draft', color: '#a389c4', deadline: '2027-01-22', progress: 8, contractNumber: 'Номер договора', contractName: 'Название договора', correspondencePath: '' },
  ],
  tasks: [
    { id: 't1', title: 'Подготовить сводную таблицу замечаний', projectId: 'p1', due: '2026-08-28', priority: 'high', done: false, description: '', status: 'planned', assignee: '', source: 'manual', contractStageId: '', meetingId: '' },
    { id: 't2', title: 'Согласовать исходные данные с заказчиком', projectId: 'p2', due: '2026-08-31', priority: 'medium', done: false, description: '', status: 'in_progress', assignee: '', source: 'manual', contractStageId: '', meetingId: '' },
    { id: 't3', title: 'Направить протокол совещания', projectId: 'p1', due: '2026-09-02', priority: 'medium', done: false, description: '', status: 'planned', assignee: '', source: 'meeting', contractStageId: '', meetingId: '' },
    { id: 't4', title: 'Проверить комплект документов', projectId: 'p3', due: '2026-09-05', priority: 'low', done: false, description: '', status: 'planned', assignee: '', source: 'manual', contractStageId: '', meetingId: '' },
    { id: 't5', title: 'Обновить календарный план', projectId: 'p2', due: '2026-08-26', priority: 'high', done: true, description: '', status: 'done', assignee: '', source: 'manual', contractStageId: '', meetingId: '' },
  ],
  notes: [
    { id: 'n1', text: 'Уточнить у заказчика состав оборудования для этапа 2.', projectId: 'p1', createdAt: '2026-08-27T09:20:00' },
    { id: 'n2', text: 'Идея: добавить шаблон еженедельного отчёта.', createdAt: '2026-08-26T15:40:00' },
  ],
  links: [
    { id: 'l1', label: 'Папка проекта', target: 'C:\\Projects\\PR-024', projectId: 'p1', kind: 'file' },
    { id: 'l2', label: 'YouTrack', target: 'https://www.jetbrains.com/youtrack/', projectId: 'p1', kind: 'web' },
    { id: 'l3', label: 'Реестр документов', target: 'C:\\Projects\\PR-031\\Documents', projectId: 'p2', kind: 'file' },
  ],
  contractStages: [
    { id: 'cs1', projectId: 'p1', number: '1', name: 'Предпроектная подготовка', scope: 'Сбор исходных данных, обследование объекта', startDate: '2026-07-01', endDate: '2026-08-15', status: 'closed', suspensionLetter: '', cost: 0, advancePercent: 0, advanceStatus: 'paid', finalPaymentStatus: 'paid', comment: '' },
    { id: 'cs2', projectId: 'p1', number: '2', name: 'Проектная документация', scope: 'Разработка основных разделов проекта', startDate: '2026-08-16', endDate: '2026-10-30', status: 'started', suspensionLetter: '', cost: 0, advancePercent: 0, advanceStatus: 'paid', finalPaymentStatus: 'unpaid', comment: 'Срок уточняется' },
  ],
  meetings: [],
  inbox: [],
  trackings: [],
  integrations: {
    worklogExcelPath: '',
    youtrackTokenPath: '',
    outlookBridgeKey: '',
  },
}
