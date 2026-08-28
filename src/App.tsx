import { useEffect, useMemo, useState } from 'react'
import { getAllWindows, getCurrentWindow } from '@tauri-apps/api/window'
import { invoke } from '@tauri-apps/api/core'
import { listen } from '@tauri-apps/api/event'
import {
  ArrowUpRight, BellRing, BriefcaseBusiness, CalendarDays, Check, CheckCircle2,
  ChevronRight, Circle, Clock3, ExternalLink, FileText, FolderOpen,
  Eye, EyeOff, GripVertical, LayoutDashboard, Link2, ListTodo, Menu, MessageSquareText, MoreHorizontal,
  Inbox, NotebookPen, PanelLeftClose, Pencil, Plus, Settings, StickyNote,
  Trash2, X,
} from 'lucide-react'
import { initialData } from './data'
import type { CapturedEmail, ContractStage, ContractStageStatus, Meeting, MeetingAction, PaymentStatus, PlannerData, Priority, Project, Task, TaskSource, TaskStatus, Tracking, TrackingStatus } from './types'

type View = 'dashboard' | 'inbox' | 'projects' | 'tasks' | 'notes' | 'project'

const navItems: { id: View; label: string; icon: typeof LayoutDashboard }[] = [
  { id: 'dashboard', label: 'Обзор', icon: LayoutDashboard },
  { id: 'inbox', label: 'Входящие', icon: Inbox },
  { id: 'projects', label: 'Проекты', icon: BriefcaseBusiness },
  { id: 'tasks', label: 'Задачи', icon: ListTodo },
  { id: 'notes', label: 'Заметки', icon: NotebookPen },
]

const priorityLabel: Record<Priority, string> = { high: 'Важно', medium: 'Обычно', low: 'Не срочно' }

async function openFileLink(target: string) {
  if ('__TAURI_INTERNALS__' in window) {
    try {
      await invoke('open_local_path', { path: target })
    } catch (error) {
      window.alert(String(error))
    }
    return
  }

  await navigator.clipboard.writeText(target)
  window.alert('Путь скопирован. Открытие локальных файлов доступно в настольной версии планера.')
}

async function openWebLink(target: string) {
  const url = safeWebUrl(target)
  if ('__TAURI_INTERNALS__' in window) {
    try {
      await invoke('open_web_url', { url })
    } catch (error) {
      window.alert(String(error))
    }
    return
  }

  window.open(url, '_blank', 'noopener,noreferrer')
}

function loadData(): PlannerData {
  try {
    const saved = localStorage.getItem('myplanner:data')
    if (!saved) return initialData
    const parsed = JSON.parse(saved) as Partial<PlannerData>
    const statusMigration: Record<string, ContractStageStatus> = { planned: 'not_started', in_progress: 'started', completed: 'work_completed', paused: 'suspended' }
    const projects = (parsed.projects || initialData.projects).map(project => ({ ...project, contractNumber: project.contractNumber || 'Номер договора', contractName: project.contractName || 'Название договора', correspondencePath: project.correspondencePath || '' }))
    const tasks = (parsed.tasks || initialData.tasks).map(task => ({ ...task, description: task.description || '', status: task.status || (task.done ? 'done' : 'planned'), assignee: task.assignee || '', source: task.source || 'manual', contractStageId: task.contractStageId || '', meetingId: task.meetingId || '' })) as Task[]
    const contractStages = (parsed.contractStages || initialData.contractStages).map(stage => {
      const legacy = stage as ContractStage & { payment?: string; advanceAmount?: number }
      return { ...stage, status: statusMigration[stage.status] || stage.status, suspensionLetter: stage.suspensionLetter || '', cost: Number(stage.cost) || 0, advancePercent: Number(stage.advancePercent ?? legacy.advanceAmount) || 0, advanceStatus: stage.advanceStatus || (legacy.payment === 'Оплачено' ? 'paid' : 'unpaid'), finalPaymentStatus: stage.finalPaymentStatus || 'unpaid' } as ContractStage
    })
    return { ...initialData, ...parsed, projects, tasks, contractStages, meetings: parsed.meetings || [], inbox: parsed.inbox || [], trackings: parsed.trackings || [], integrations: { ...initialData.integrations, ...(parsed.integrations || {}) } }
  } catch { return initialData }
}

function uid(prefix: string) {
  return `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`
}

function App() {
  const [data, setData] = useState<PlannerData>(loadData)
  const [view, setView] = useState<View>('dashboard')
  const [selectedProjectId, setSelectedProjectId] = useState(data.projects[0]?.id || '')
  const nativeSticker = new URLSearchParams(window.location.search).get('mode') === 'sticky'
  const [compact, setCompact] = useState(nativeSticker)
  const [sidebar, setSidebar] = useState(true)
  const [quickOpen, setQuickOpen] = useState(false)
  const [settingsOpen, setSettingsOpen] = useState(false)

  useEffect(() => localStorage.setItem('myplanner:data', JSON.stringify(data)), [data])
  useEffect(() => {
    const syncData = (event: StorageEvent) => {
      if (event.key !== 'myplanner:data' || !event.newValue) return
      try { setData(JSON.parse(event.newValue) as PlannerData) } catch { /* ignore malformed external data */ }
    }
    window.addEventListener('storage', syncData)
    return () => window.removeEventListener('storage', syncData)
  }, [])
  useEffect(() => {
    if (!('__TAURI_INTERNALS__' in window)) return
    let disposed = false
    let unlisten: (() => void) | undefined
    const connect = async () => {
      let key = data.integrations.outlookBridgeKey
      if (!key) {
        key = crypto.randomUUID().replaceAll('-', '')
        setData(current => ({ ...current, integrations: { ...current.integrations, outlookBridgeKey: key } }))
      }
      await invoke('set_outlook_bridge_key', { key })
      unlisten = await listen<CapturedEmail>('outlook-email-captured', event => {
        setData(current => {
          if (current.inbox.some(item => item.outlookItemId && item.outlookItemId === event.payload.outlookItemId)) return current
          return { ...current, inbox: [event.payload, ...current.inbox] }
        })
      })
      if (disposed) unlisten?.()
    }
    void connect()
    return () => { disposed = true; unlisten?.() }
  }, [data.integrations.outlookBridgeKey])

  const openTasks = data.tasks.filter(t => !t.done)
  const selectedProject = data.projects.find(project => project.id === selectedProjectId)

  const toggleTask = (id: string) => setData(d => ({
    ...d, tasks: d.tasks.map(t => t.id === id ? { ...t, done: !t.done, status: !t.done ? 'done' : 'planned' } : t),
  }))

  const expandPlanner = async () => {
    if (nativeSticker && '__TAURI_INTERNALS__' in window) {
      const windows = await getAllWindows()
      const main = windows.find(item => item.label === 'main')
      await main?.show()
      await main?.unminimize()
      await main?.setFocus()
      return
    }
    setCompact(false)
  }

  const hideSticker = async () => {
    if (nativeSticker && '__TAURI_INTERNALS__' in window) await getCurrentWindow().hide()
  }

  if (compact) return <CompactView data={data} setData={setData} onExpand={expandPlanner} onHide={nativeSticker ? hideSticker : undefined} native={nativeSticker} />

  return (
    <div className={`app-shell ${sidebar ? '' : 'sidebar-hidden'}`}>
      <aside className="sidebar">
        <div className="brand"><div className="brand-mark"><Check /></div><span>myplanner</span></div>
        <nav>
          <p className="nav-title">Рабочее пространство</p>
          {navItems.map(item => <button key={item.id} className={view === item.id ? 'active' : ''} onClick={() => setView(item.id)}><item.icon /><span>{item.label}</span>{item.id === 'tasks' && <b>{openTasks.length}</b>}{item.id === 'inbox' && data.inbox.filter(email => !email.processed).length > 0 && <b>{data.inbox.filter(email => !email.processed).length}</b>}</button>)}
          <p className="nav-title projects-title">Проекты</p>
          {data.projects.map(project => <button key={project.id} className={`project-nav ${view === 'project' && selectedProjectId === project.id ? 'active' : ''}`} onClick={() => { setSelectedProjectId(project.id); setView('project') }}><i style={{ background: project.color }} /> <span>{project.name}</span></button>)}
        </nav>
        <div className="sidebar-bottom">
          <button onClick={() => setCompact(true)}><StickyNote /><span>Режим стикера</span><ArrowUpRight /></button>
          <button onClick={() => setSettingsOpen(true)}><Settings /><span>Настройки</span></button>
          <div className="profile"><div>Д</div><span><strong>Денис</strong><small>Моё пространство</small></span><MoreHorizontal /></div>
        </div>
      </aside>

      <main>
        <header>
          <button className="icon-button menu-button" onClick={() => setSidebar(v => !v)}>{sidebar ? <PanelLeftClose /> : <Menu />}</button>
          {view === 'project' && selectedProject && <div className="header-project"><strong>{selectedProject.name}</strong><span className={`status ${selectedProject.status}`}>{selectedProject.status === 'active' ? 'В работе' : selectedProject.status === 'paused' ? 'На паузе' : 'Черновик'}</span></div>}
        </header>

        <section className="content">
          {view === 'dashboard' && <Dashboard data={data} setData={setData} toggleTask={toggleTask} onAdd={() => setQuickOpen(true)} />}
          {view === 'inbox' && <InboxView data={data} setData={setData} />}
          {view === 'projects' && <Projects data={data} setData={setData} />}
          {view === 'tasks' && <Tasks data={data} tasks={data.tasks} setData={setData} toggleTask={toggleTask} />}
          {view === 'notes' && <Notes data={data} setData={setData} />}
          {view === 'project' && <ProjectWorkspace project={data.projects.find(project => project.id === selectedProjectId)} data={data} setData={setData} />}
        </section>
      </main>
      {quickOpen && <QuickAdd data={data} setData={setData} close={() => setQuickOpen(false)} />}
      {settingsOpen && <DataSettings data={data} importData={setData} close={() => setSettingsOpen(false)} />}
    </div>
  )
}

function DataSettings({ data, importData, close }: { data: PlannerData; importData: React.Dispatch<React.SetStateAction<PlannerData>>; close: () => void }) {
  const [message, setMessage] = useState('')
  const exportData = () => {
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' })
    const url = URL.createObjectURL(blob)
    const anchor = document.createElement('a')
    anchor.href = url
    anchor.download = `myplanner-backup-${new Date().toISOString().slice(0, 10)}.json`
    anchor.click()
    URL.revokeObjectURL(url)
    setMessage('Резервная копия сохранена.')
  }
  const importFile = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) return
    try {
      const parsed = JSON.parse(await file.text()) as Partial<PlannerData>
      if (!Array.isArray(parsed.projects) || !Array.isArray(parsed.tasks) || !Array.isArray(parsed.notes) || !Array.isArray(parsed.links)) throw new Error('invalid data')
      const normalized: PlannerData = {
        ...initialData,
        ...parsed,
        projects: parsed.projects.map(project => ({ ...project, contractNumber: project.contractNumber || 'Номер договора', contractName: project.contractName || 'Название договора', correspondencePath: project.correspondencePath || '' })),
        tasks: parsed.tasks.map(task => ({ ...task, description: task.description || '', status: task.status || (task.done ? 'done' : 'planned'), assignee: task.assignee || '', source: task.source || 'manual', contractStageId: task.contractStageId || '', meetingId: task.meetingId || '' })) as Task[],
        contractStages: parsed.contractStages || [],
        meetings: parsed.meetings || [],
        inbox: parsed.inbox || [],
        trackings: parsed.trackings || [],
        integrations: { ...initialData.integrations, ...(parsed.integrations || {}) },
      }
      importData(normalized)
      setMessage(`Импортировано: ${normalized.projects.length} проектов, ${normalized.tasks.length} задач.`)
    } catch {
      setMessage('Не удалось импортировать файл: неверный формат данных.')
    } finally { event.target.value = '' }
  }
  return <div className="modal-backdrop" onMouseDown={event => event.target === event.currentTarget && close()}><div className="modal data-settings"><div className="modal-head"><div><span>Настройки</span><h2>Данные MyPlanner</h2></div><button onClick={close}><X /></button></div><p className="settings-intro">Создайте резервную копию или перенесите данные между браузером и приложением.</p><div className="data-action"><div><strong>Экспорт данных</strong><p>Сохранить все проекты, задачи, ссылки, договоры и совещания в JSON-файл.</p></div><button className="quiet-button" onClick={exportData}>Экспортировать</button></div><div className="data-action"><div><strong>Импорт данных</strong><p>Заменить текущие данные содержимым ранее сохранённой резервной копии.</p></div><label className="primary-button file-button">Выбрать файл<input type="file" accept="application/json,.json" onChange={importFile} /></label></div>{message && <div className="settings-message">{message}</div>}<div className="modal-actions"><button className="quiet-button" onClick={close}>Закрыть</button></div></div></div>
}

function InboxView({ data, setData }: { data: PlannerData; setData: React.Dispatch<React.SetStateAction<PlannerData>> }) {
  const [projectByEmail, setProjectByEmail] = useState<Record<string, string>>({})
  const [showSetup, setShowSetup] = useState(false)
  const pending = data.inbox.filter(email => !email.processed)
  const emailText = (email: CapturedEmail) => [email.subject, `От: ${email.senderName || email.senderEmail}${email.senderEmail && email.senderName ? ` <${email.senderEmail}>` : ''}`, email.receivedAt ? `Дата: ${new Date(email.receivedAt).toLocaleString('ru-RU')}` : '', email.excerpt].filter(Boolean).join('\n\n')
  const finish = (email: CapturedEmail, kind: 'note' | 'task' | 'tracking') => {
    const projectId = projectByEmail[email.id] || ''
    if (kind === 'tracking' && !projectId) {
      window.alert('Для отслеживания сначала выберите проект.')
      return
    }
    const sentDate = email.receivedAt ? email.receivedAt.slice(0, 10) : new Date().toISOString().slice(0, 10)
    const checkDateValue = new Date(`${sentDate}T00:00:00`)
    checkDateValue.setDate(checkDateValue.getDate() + 7)
    setData(current => ({
      ...current,
      inbox: current.inbox.map(item => item.id === email.id ? { ...item, processed: true } : item),
      notes: kind === 'note' ? [{ id: uid('n'), text: emailText(email), projectId: projectId || undefined, createdAt: new Date().toISOString() }, ...current.notes] : current.notes,
      tasks: kind === 'task' ? [{ id: uid('t'), title: email.subject || 'Письмо Outlook', projectId, due: new Date().toISOString().slice(0, 10), priority: 'medium', done: false, description: emailText(email), status: 'planned', assignee: '', source: 'outlook', contractStageId: '', meetingId: '' }, ...current.tasks] : current.tasks,
      trackings: kind === 'tracking' ? [{ id: uid('tr'), projectId, subject: email.subject || 'Запрос из Outlook', recipient: email.recipients || '', sentDate, checkDate: checkDateValue.toISOString().slice(0, 10), status: 'waiting', comment: emailText(email), sourceEmailId: email.outlookItemId }, ...current.trackings] : current.trackings,
    }))
  }
  const remove = (id: string) => setData(current => ({ ...current, inbox: current.inbox.filter(email => email.id !== id) }))
  const copyKey = async () => {
    await navigator.clipboard.writeText(data.integrations.outlookBridgeKey)
    window.alert('Код сопряжения скопирован.')
  }
  return <>
    <PageIntro eyebrow="Outlook" title="Входящие" action={<button className="quiet-button" onClick={() => setShowSetup(true)}><Settings /> Настройка Outlook</button>} />
    {pending.length === 0 ? <div className="empty-inbox"><Inbox /><h2>Новых писем нет</h2><p>Откройте письмо в Outlook и нажмите «В MyPlanner».</p></div> : <div className="inbox-list">{pending.map(email => <article className="inbox-card" key={email.id}>
      <div className="inbox-card-head"><div><span>{email.senderName || email.senderEmail || 'Отправитель не указан'}</span><h2>{email.subject || 'Без темы'}</h2></div><time>{email.receivedAt ? new Date(email.receivedAt).toLocaleString('ru-RU') : ''}</time></div>
      {email.excerpt && <p>{email.excerpt}</p>}
      <div className="inbox-actions"><select value={projectByEmail[email.id] || ''} onChange={event => setProjectByEmail(current => ({ ...current, [email.id]: event.target.value }))}><option value="">Без проекта</option>{data.projects.map(project => <option key={project.id} value={project.id}>{project.name}</option>)}</select><button className="quiet-button" onClick={() => finish(email, 'note')}><NotebookPen /> В заметку</button><button className="quiet-button" onClick={() => finish(email, 'tracking')}><BellRing /> Отслеживать</button><button className="primary-button" onClick={() => finish(email, 'task')}><ListTodo /> В задачу</button><button className="icon-button" title="Удалить" onClick={() => remove(email.id)}><Trash2 /></button></div>
    </article>)}</div>}
    {showSetup && <div className="modal-backdrop" onMouseDown={event => event.target === event.currentTarget && setShowSetup(false)}><div className="modal outlook-setup"><div className="modal-head"><div><span>Интеграция</span><h2>Outlook Web Add-in</h2></div><button onClick={() => setShowSetup(false)}><X /></button></div><p className="settings-intro">Введите этот код один раз в надстройке Outlook. MyPlanner должен быть запущен во время передачи письма.</p><label>Код сопряжения<div className="pairing-row"><input readOnly value={data.integrations.outlookBridgeKey} /><button className="quiet-button" onClick={copyKey}>Копировать</button></div></label><small>Локальный приёмник: http://127.0.0.1:17832</small><div className="modal-actions"><button className="quiet-button" onClick={() => setShowSetup(false)}>Закрыть</button></div></div></div>}
  </>
}

function PageIntro({ eyebrow, title, subtitle, action }: { eyebrow: string; title: string; subtitle?: string; action?: React.ReactNode }) {
  return <div className="page-intro"><div><span>{eyebrow}</span><h1>{title}</h1>{subtitle && <p>{subtitle}</p>}</div>{action}</div>
}

function Dashboard({ data, setData, toggleTask, onAdd }: { data: PlannerData; setData: React.Dispatch<React.SetStateAction<PlannerData>>; toggleTask: (id: string) => void; onAdd: () => void }) {
  const today = new Intl.DateTimeFormat('ru-RU', { weekday: 'long', day: 'numeric', month: 'long' }).format(new Date())
  const open = data.tasks.filter(t => !t.done)
  return <>
    <PageIntro eyebrow={today} title="Добрый вечер, Денис" action={<button className="primary-button" onClick={onAdd}><Plus /> Быстрая запись</button>} />
    <div className="stats-grid">
      <div className="stat-card coral"><span><BriefcaseBusiness /> Активные проекты</span><strong>{data.projects.filter(p => p.status === 'active').length}</strong><small>из {data.projects.length} проектов</small></div>
      <div className="stat-card green"><span><CheckCircle2 /> Открытые задачи</span><strong>{open.length}</strong><small><b>{open.filter(t => t.priority === 'high').length}</b> требуют внимания</small></div>
      <div className="stat-card violet"><span><Clock3 /> Ближайший срок</span><strong>28 авг</strong><small>остался 1 день</small></div>
      <div className="stat-card sand"><span><FileText /> Документы</span><strong>12</strong><small>2 ожидают ответа</small></div>
    </div>
    <div className="dashboard-grid">
      <div className="panel task-panel">
        <div className="panel-head"><div><h2>Ближайшие задачи</h2><p>Сосредоточьтесь на главном</p></div><button>Все задачи <ChevronRight /></button></div>
        <div className="task-list">
          {open.slice(0, 4).map(task => <TaskRow key={task.id} task={task} project={data.projects.find(p => p.id === task.projectId)} toggle={() => toggleTask(task.id)} />)}
        </div>
      </div>
    </div>
    <div className="bottom-grid">
      <div className="panel notes-preview"><div className="panel-head"><div><h2>Последние заметки</h2><p>Мысли и важные детали</p></div><StickyNote /></div>{data.notes.slice(0,2).map(note => <article key={note.id}><p>{note.text}</p><small>{new Date(note.createdAt).toLocaleDateString('ru-RU')} · {data.projects.find(p => p.id === note.projectId)?.name || 'Без проекта'}</small></article>)}</div>
    </div>
    <WorklogTransfer data={data} setData={setData} />
  </>
}

interface WorklogPreviewEntry {
  issueKey: string
  minutes: number
  workDate: string
}

interface WorklogPreview {
  weekStart: string
  weekEnd: string
  workDate?: string
  entries: WorklogPreviewEntry[]
  totalMinutes: number
  errors: string[]
}

interface WorklogSendResult {
  created: number
  skipped: number
  failed: number
  reportPath?: string
  items: { issueKey: string; status: string; message: string }[]
}

function displayIsoDate(value?: string) {
  return value ? new Date(`${value}T00:00:00`).toLocaleDateString('ru-RU') : '—'
}

function displayMinutes(minutes: number) {
  const hours = Math.floor(minutes / 60)
  const remainder = minutes % 60
  return [hours > 0 ? `${hours} ч` : '', remainder > 0 ? `${remainder} мин` : ''].filter(Boolean).join(' ') || '0 мин'
}

function WorklogTransfer({ data, setData }: { data: PlannerData; setData: React.Dispatch<React.SetStateAction<PlannerData>> }) {
  const [selectedDate, setSelectedDate] = useState(new Date().toISOString().slice(0, 10))
  const [preview, setPreview] = useState<WorklogPreview | null>(null)
  const [result, setResult] = useState<WorklogSendResult | null>(null)
  const [busy, setBusy] = useState<'preview' | 'send' | ''>('')
  const [error, setError] = useState('')
  const [settingsOpen, setSettingsOpen] = useState(false)
  const settings = data.integrations
  const updateSettings = (patch: Partial<PlannerData['integrations']>) => {
    setData(current => ({ ...current, integrations: { ...current.integrations, ...patch } }))
    setPreview(null)
    setResult(null)
  }
  const chooseFile = async (command: 'choose_worklog_excel' | 'choose_youtrack_token_file', key: keyof PlannerData['integrations']) => {
    if (!('__TAURI_INTERNALS__' in window)) { setError('Работа с Excel и YouTrack доступна в настольной версии.'); return }
    const path = await invoke<string | null>(command)
    if (path) updateSettings({ [key]: path })
  }
  const checkWeek = async () => {
    setBusy('preview'); setError(''); setResult(null); setPreview(null)
    try {
      setPreview(await invoke<WorklogPreview>('preview_youtrack_week', { excelPath: settings.worklogExcelPath, selectedDate }))
    } catch (previewError) {
      setError(String(previewError))
    } finally { setBusy('') }
  }
  const send = async () => {
    if (!preview || preview.errors.length || !preview.entries.length) return
    if (!window.confirm(`Создать ${preview.entries.length} записей на ${displayIsoDate(preview.workDate)} общей длительностью ${displayMinutes(preview.totalMinutes)}?`)) return
    setBusy('send'); setError(''); setResult(null)
    try {
      setResult(await invoke<WorklogSendResult>('send_youtrack_worklogs', { excelPath: settings.worklogExcelPath, tokenPath: settings.youtrackTokenPath, entries: preview.entries }))
    } catch (sendError) {
      setError(String(sendError))
    } finally { setBusy('') }
  }
  const canPreview = Boolean(settings.worklogExcelPath && selectedDate && !busy)
  const canSend = Boolean(preview && !preview.errors.length && preview.entries.length && settings.youtrackTokenPath && !busy)
  return <div className="panel worklog-panel"><div className="panel-head worklog-panel-head"><div><h2>Фиксация времени в YouTrack</h2></div><button className="worklog-settings-button" onClick={() => setSettingsOpen(true)} title="Настройки файлов"><Settings /></button></div><div className="worklog-controls"><input type="date" aria-label="Дата нужной недели" value={selectedDate} onChange={event => { setSelectedDate(event.target.value); setPreview(null); setResult(null) }} /><button className="quiet-button" onClick={() => void checkWeek()} disabled={!canPreview}>{busy === 'preview' ? 'Читаю Excel…' : 'Проверить неделю'}</button></div>{busy && <div className="worklog-progress"><i /><span>{busy === 'preview' ? 'Анализ выбранной недели…' : 'Проверка дублей и отправка…'}</span></div>}{error && <div className="correspondence-error">{error}</div>}{preview && <div className="worklog-preview"><div className="worklog-week"><span>Неделя {displayIsoDate(preview.weekStart)} — {displayIsoDate(preview.weekEnd)}</span><strong>{preview.entries.length} задач · {displayMinutes(preview.totalMinutes)}</strong></div>{preview.errors.length > 0 && <div className="worklog-errors"><strong>Предварительная проверка содержит ошибки</strong>{preview.errors.map(message => <span key={message}>{message}</span>)}</div>}<div className="worklog-table-wrap"><table><thead><tr><th>Ключ задачи</th><th>Длительность</th><th>Минуты</th><th>Дата фиксации</th></tr></thead><tbody>{preview.entries.map(entry => <tr key={entry.issueKey}><td><strong>{entry.issueKey}</strong></td><td>{displayMinutes(entry.minutes)}</td><td>{entry.minutes}</td><td>{displayIsoDate(entry.workDate)}</td></tr>)}</tbody></table></div><div className="worklog-submit"><span>Отправка не выполняется автоматически.</span><button className="primary-button" disabled={!canSend} onClick={() => void send()}>Отправить в YouTrack</button></div></div>}{result && <div className="worklog-result"><div><span>Создано<strong>{result.created}</strong></span><span>Пропущено<strong>{result.skipped}</strong></span><span>Ошибок<strong>{result.failed}</strong></span></div>{result.items.some(item => item.status === 'failed') && <ul>{result.items.filter(item => item.status === 'failed').map(item => <li key={item.issueKey}><b>{item.issueKey}</b>: {item.message}</li>)}</ul>}{result.reportPath && <p>JSON-отчёт: <button onClick={() => openFileLink(result.reportPath!)}>{result.reportPath}</button></p>}</div>}{settingsOpen && <div className="modal-backdrop" onMouseDown={event => event.target === event.currentTarget && setSettingsOpen(false)}><div className="modal worklog-settings-modal"><div className="modal-head"><div><span>YouTrack</span><h2>Настройки файлов</h2></div><button onClick={() => setSettingsOpen(false)}><X /></button></div><div className="data-action"><div><strong>Excel с трудозатратами</strong><p>{settings.worklogExcelPath || 'Файл не выбран'}</p></div><button className="quiet-button" onClick={() => void chooseFile('choose_worklog_excel', 'worklogExcelPath')}>{settings.worklogExcelPath ? 'Изменить' : 'Выбрать'}</button></div><div className="data-action"><div><strong>Файл токена YouTrack</strong><p>{settings.youtrackTokenPath || 'Файл не выбран'}</p></div><button className="quiet-button" onClick={() => void chooseFile('choose_youtrack_token_file', 'youtrackTokenPath')}>{settings.youtrackTokenPath ? 'Изменить' : 'Выбрать'}</button></div><div className="modal-actions"><button className="primary-button" onClick={() => setSettingsOpen(false)}>Готово</button></div></div></div>}</div>
}

function TaskRow({ task, project, toggle }: { task: PlannerData['tasks'][number]; project?: Project; toggle: () => void }) {
  return <div className={`task-row ${task.done ? 'done' : ''}`}><button className="check-button" onClick={toggle}>{task.done ? <CheckCircle2 /> : <Circle />}</button><div className="task-copy"><strong>{task.title}</strong><span><i style={{ background: project?.color }} />{project?.name}</span></div><span className={`priority ${task.priority}`}>{priorityLabel[task.priority]}</span><time><CalendarDays />{new Date(task.due).toLocaleDateString('ru-RU', { day: 'numeric', month: 'short' })}</time></div>
}

function Projects({ data, setData }: { data: PlannerData; setData: React.Dispatch<React.SetStateAction<PlannerData>> }) {
  const [editing, setEditing] = useState<Project | null>(null)
  const addProject = () => setData(d => ({ ...d, projects: [...d.projects, { id: uid('p'), code: `PR-${String(d.projects.length + 37).padStart(3, '0')}`, name: 'Новый проект', client: 'Заказчик не указан', status: 'draft', color: '#d19a54', deadline: '2026-12-31', progress: 0, contractNumber: 'Номер договора', contractName: 'Название договора', correspondencePath: '' }] }))
  return <><PageIntro eyebrow="Портфель" title="Проекты" subtitle="Все рабочие направления в одном месте." action={<button className="primary-button" onClick={addProject}><Plus /> Новый проект</button>} /><div className="project-cards">{data.projects.map(p => <article key={p.id} className="project-card"><div className="project-card-top"><i className="project-color-dot" style={{ background: p.color }} title="Цвет проекта" /><button className="edit-project" onClick={() => setEditing(p)} title="Редактировать проект"><Pencil /> Редактировать</button></div><div className="project-title-row"><h2>{p.name}</h2><span className={`status ${p.status}`}>{p.status === 'active' ? 'В работе' : p.status === 'paused' ? 'На паузе' : 'Черновик'}</span></div><p>{p.client}</p><span className="project-id"><b>{p.code || 'Не указан'}</b></span></article>)}</div>{editing && <ProjectEditor project={editing} close={() => setEditing(null)} save={project => { setData(d => ({ ...d, projects: d.projects.map(p => p.id === project.id ? project : p) })); setEditing(null) }} />}</>
}

function ProjectEditor({ project, save, close }: { project: Project; save: (project: Project) => void; close: () => void }) {
  const [form, setForm] = useState<Project>({ ...project })
  const update = <K extends keyof Project>(key: K, value: Project[K]) => setForm(current => ({ ...current, [key]: value }))
  const submit = (event: React.FormEvent) => {
    event.preventDefault()
    if (!form.name.trim()) return
    save({ ...form, name: form.name.trim(), code: form.code.trim(), client: form.client.trim() || 'Заказчик не указан' })
  }
  return <div className="modal-backdrop" onMouseDown={event => event.target === event.currentTarget && close()}>
    <form className="modal project-editor" onSubmit={submit}>
      <div className="modal-head"><div><span>Базовые настройки</span><h2>Редактирование проекта</h2></div><button type="button" onClick={close}><X /></button></div>
      <div className="form-grid">
        <label className="wide">Название проекта<input autoFocus value={form.name} onChange={e => update('name', e.target.value)} required /></label>
        <label>ID проекта<input value={form.code} onChange={e => update('code', e.target.value)} placeholder="Например, PR-024" /></label>
        <label>Статус<select value={form.status} onChange={e => update('status', e.target.value as Project['status'])}><option value="draft">Черновик</option><option value="active">В работе</option><option value="paused">На паузе</option></select></label>
        <label>Заказчик<input value={form.client} onChange={e => update('client', e.target.value)} /></label>
        <label className="wide">Цвет проекта<div className="color-field"><input type="color" value={form.color} onChange={e => update('color', e.target.value)} /><span>{form.color.toUpperCase()}</span></div></label>
      </div>
      <div className="modal-actions"><button type="button" className="quiet-button" onClick={close}>Отмена</button><button type="submit" className="primary-button">Сохранить изменения</button></div>
    </form>
  </div>
}

type ProjectSection = 'links' | 'contract' | 'correspondence' | 'tasks' | 'trackings' | 'meetings'

function ProjectWorkspace({ project, data, setData }: { project?: Project; data: PlannerData; setData: React.Dispatch<React.SetStateAction<PlannerData>> }) {
  const [section, setSection] = useState<ProjectSection>('tasks')
  const [addingLink, setAddingLink] = useState(false)
  const [editingLink, setEditingLink] = useState<PlannerData['links'][number] | null>(null)
  const [draggingLinkId, setDraggingLinkId] = useState('')
  const [dropTargetId, setDropTargetId] = useState('')
  const [addingStage, setAddingStage] = useState(false)
  const [editingStage, setEditingStage] = useState<ContractStage | null>(null)
  if (!project) return <div className="empty-project"><FolderOpen /><h2>Проект не найден</h2><p>Выберите проект в боковом меню.</p></div>
  const projectLinks = data.links.filter(link => link.projectId === project.id)
  const activeTrackings = data.trackings.filter(tracking => tracking.projectId === project.id && tracking.status !== 'completed' && tracking.status !== 'cancelled')
  const sections: { id: ProjectSection; label: string; icon: typeof Link2 }[] = [
    { id: 'tasks', label: 'Задачи', icon: ListTodo },
    { id: 'trackings', label: 'Отслеживания', icon: BellRing },
    { id: 'links', label: 'Ссылки', icon: Link2 },
    { id: 'meetings', label: 'Совещания', icon: CalendarDays },
    { id: 'correspondence', label: 'Переписка', icon: MessageSquareText },
    { id: 'contract', label: 'Договор', icon: FileText },
  ]
  return <div className="project-workspace">
    <div className="workspace-tabs" role="tablist">
      {sections.map(item => <button key={item.id} className={section === item.id ? 'active' : ''} onClick={() => setSection(item.id)}><item.icon />{item.label}{item.id === 'links' && projectLinks.length > 0 && <b>{projectLinks.length}</b>}{item.id === 'tasks' && data.tasks.filter(task => task.projectId === project.id && !task.done).length > 0 && <b>{data.tasks.filter(task => task.projectId === project.id && !task.done).length}</b>}{item.id === 'trackings' && activeTrackings.length > 0 && <b>{activeTrackings.length}</b>}</button>)}
    </div>
    <div className="workspace-body">
      {section === 'links' && <div className="workspace-section">
        <div className="section-heading"><div><span>Материалы проекта</span><h2>Ссылки</h2><p>Веб-ресурсы, сетевые папки и файлы проекта.</p></div><button className="primary-button" onClick={() => setAddingLink(true)}><Plus /> Добавить ссылку</button></div>
        {projectLinks.length > 0 ? <div className="link-preview-list">{projectLinks.map(link => <div key={link.id} className={`${draggingLinkId === link.id ? 'dragging' : ''} ${dropTargetId === link.id && draggingLinkId !== link.id ? 'drop-target' : ''}`} onDragOver={event => { event.preventDefault(); setDropTargetId(link.id) }} onDragLeave={() => setDropTargetId(current => current === link.id ? '' : current)} onDrop={event => { event.preventDefault(); if (draggingLinkId && draggingLinkId !== link.id) setData(current => reorderProjectLinks(current, project.id, draggingLinkId, link.id)); setDraggingLinkId(''); setDropTargetId('') }}><div><button className="drag-handle" draggable title="Перетащить ссылку" onDragStart={event => { setDraggingLinkId(link.id); event.dataTransfer.effectAllowed = 'move'; event.dataTransfer.setData('text/plain', link.id) }} onDragEnd={() => { setDraggingLinkId(''); setDropTargetId('') }}><GripVertical /></button><span>{link.kind === 'file' ? <FolderOpen /> : <ExternalLink />}</span><div><strong>{link.label}</strong><small>{link.target}</small></div></div><div className="link-controls"><button className="letter-open-button" title="Открыть" onClick={() => link.kind === 'web' ? openWebLink(link.target) : openFileLink(link.target)}>{link.kind === 'file' ? <FolderOpen /> : <ExternalLink />}</button><div className="table-actions"><button title="Редактировать" onClick={() => setEditingLink(link)}><Pencil /></button><button className="danger" title="Удалить" onClick={() => { if (window.confirm(`Удалить ссылку «${link.label}»?`)) setData(current => ({ ...current, links: current.links.filter(item => item.id !== link.id) })) }}><Trash2 /></button></div></div></div>)}</div> : <SectionPlaceholder icon={Link2} title="Ссылок пока нет" text="Добавьте веб-ссылку либо путь к локальному файлу или папке." />}
      </div>}
      {section === 'contract' && (() => {
        const projectStages = data.contractStages.filter(stage => stage.projectId === project.id)
        const totalCost = projectStages.reduce((sum, stage) => sum + (Number(stage.cost) || 0), 0)
        return <div className="workspace-section contract-section"><div className="section-heading contract-heading"><div><input className="contract-number-input" value={project.contractNumber} onChange={event => setData(current => ({ ...current, projects: current.projects.map(item => item.id === project.id ? { ...item, contractNumber: event.target.value } : item) }))} aria-label="Номер договора" placeholder="Номер договора" /><textarea className="contract-name-input" rows={1} value={project.contractName} onChange={event => setData(current => ({ ...current, projects: current.projects.map(item => item.id === project.id ? { ...item, contractName: event.target.value } : item) }))} aria-label="Название договора" placeholder="Название договора" /></div><div className="contract-total"><span>Стоимость договора</span><strong>{formatMoney(totalCost)}</strong></div><button className="primary-button" onClick={() => setAddingStage(true)}><Plus /> Добавить этап</button></div><ContractTable stages={projectStages} edit={setEditingStage} remove={stage => { if (window.confirm(`Удалить этап «${stage.name}»?`)) setData(current => ({ ...current, contractStages: current.contractStages.filter(item => item.id !== stage.id) })) }} /></div>
      })()}
      {section === 'correspondence' && <CorrespondenceSection project={project} setData={setData} />}
      {section === 'tasks' && <ProjectTasks project={project} data={data} setData={setData} />}
      {section === 'trackings' && <ProjectTrackings project={project} data={data} setData={setData} />}
      {section === 'meetings' && <ProjectMeetings project={project} data={data} setData={setData} />}
    </div>
    {addingLink && <LinkEditor projectId={project.id} close={() => setAddingLink(false)} save={link => { setData(current => ({ ...current, links: [...current.links, link] })); setAddingLink(false) }} />}
    {editingLink && <LinkEditor projectId={project.id} link={editingLink} close={() => setEditingLink(null)} save={link => { setData(current => ({ ...current, links: current.links.map(item => item.id === link.id ? link : item) })); setEditingLink(null) }} />}
    {addingStage && <ContractStageEditor projectId={project.id} close={() => setAddingStage(false)} save={stage => { setData(current => ({ ...current, contractStages: [...current.contractStages, stage] })); setAddingStage(false) }} />}
    {editingStage && <ContractStageEditor projectId={project.id} stage={editingStage} close={() => setEditingStage(null)} save={stage => { setData(current => ({ ...current, contractStages: current.contractStages.map(item => item.id === stage.id ? stage : item) })); setEditingStage(null) }} />}
  </div>
}

interface CorrespondenceFolderRecord {
  name: string
  path: string
}

interface ParsedCorrespondence extends CorrespondenceFolderRecord {
  date: string
  number: string
  direction: 'incoming' | 'outgoing' | 'unknown'
  correspondent: string
  subject: string
  draft: boolean
}

function parseCorrespondenceFolder(folder: CorrespondenceFolderRecord): ParsedCorrespondence {
  const match = folder.name.match(/^\[(\d{4})\.(\d{2})\.(\d{2}|xx)\]\s+\[([^\]]+)\]\s+(.+?)\s+-\s+(.+?)\.\s+(.+)$/i)
  if (!match) return { ...folder, date: '—', number: '—', direction: 'unknown', correspondent: 'Формат не распознан', subject: folder.name, draft: false }
  const [, year, month, day, number, from, to, subject = ''] = match
  const fromLabs = from.trim().toUpperCase() === 'ЛАБС'
  const toLabs = to.trim().toUpperCase() === 'ЛАБС'
  return {
    ...folder,
    date: `${day.toLowerCase() === 'xx' ? 'xx' : day}.${month}.${year}`,
    number,
    direction: fromLabs ? 'outgoing' : toLabs ? 'incoming' : 'unknown',
    correspondent: fromLabs ? to.trim() : from.trim(),
    subject: subject.trim() || 'Без темы',
    draft: day.toLowerCase() === 'xx' || /xxx/i.test(number),
  }
}

function CorrespondenceSection({ project, setData }: { project: Project; setData: React.Dispatch<React.SetStateAction<PlannerData>> }) {
  const [folders, setFolders] = useState<CorrespondenceFolderRecord[]>([])
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')
  const [addingDraft, setAddingDraft] = useState(false)
  const [editingLetter, setEditingLetter] = useState<ParsedCorrespondence | null>(null)

  const scan = async () => {
    if (!project.correspondencePath) { setFolders([]); return }
    if (!('__TAURI_INTERNALS__' in window)) { setError('Сканирование локальных папок доступно в настольной версии.'); return }
    setLoading(true)
    setError('')
    try {
      setFolders(await invoke<CorrespondenceFolderRecord[]>('scan_correspondence', { path: project.correspondencePath }))
    } catch (scanError) {
      setError(String(scanError))
      setFolders([])
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => { void scan() }, [project.correspondencePath])

  useEffect(() => {
    if (!('__TAURI_INTERNALS__' in window)) return
    let unlisten: (() => void) | undefined
    const clearDropTarget = () => document.querySelectorAll('.letter-drop-target').forEach(element => element.classList.remove('letter-drop-target'))
    const rowAt = (x: number, y: number) => document.elementFromPoint(x, y)?.closest<HTMLTableRowElement>('tr[data-letter-path]')
    void (async () => {
      const appWindow = getCurrentWindow()
      const scaleFactor = await appWindow.scaleFactor()
      unlisten = await appWindow.onDragDropEvent(async event => {
        if (event.payload.type === 'leave') { clearDropTarget(); return }
        const position = event.payload.position.toLogical(scaleFactor)
        const row = rowAt(position.x, position.y)
        clearDropTarget()
        if (row && event.payload.type !== 'drop') row.classList.add('letter-drop-target')
        if (event.payload.type !== 'drop' || !row) return
        const paths = event.payload.paths
        if (paths.length !== 1) { setError('Перетащите на строку один файл письма.'); return }
        setError('')
        try {
          const updated = await invoke<CorrespondenceFolderRecord>('attach_outgoing_letter', { folderPath: row.dataset.letterPath, filePath: paths[0] })
          setFolders(current => current.map(folder => folder.path === row.dataset.letterPath ? updated : folder).sort((left, right) => right.name.localeCompare(left.name)))
        } catch (attachError) {
          setError(String(attachError))
        }
      })
    })()
    return () => { clearDropTarget(); unlisten?.() }
  }, [project.correspondencePath])

  const chooseFolder = async () => {
    if (!('__TAURI_INTERNALS__' in window)) { setError('Выбор локальной папки доступен в настольной версии.'); return }
    const path = await invoke<string | null>('choose_correspondence_folder')
    if (!path) return
    setData(current => ({ ...current, projects: current.projects.map(item => item.id === project.id ? { ...item, correspondencePath: path } : item) }))
  }

  const deleteLetter = async (letter: ParsedCorrespondence) => {
    if (!window.confirm(`Переместить папку письма «${letter.name}» в Корзину?`)) return
    setError('')
    try {
      await invoke('delete_correspondence_folder', { folderPath: letter.path })
      setFolders(current => current.filter(folder => folder.path !== letter.path))
    } catch (deleteError) {
      setError(String(deleteError))
    }
  }

  const letters = folders.map(parseCorrespondenceFolder)
  return <div className="workspace-section correspondence-section">
    <div className="section-heading correspondence-heading"><div><h2>Переписка</h2><div className="correspondence-path"><span>{project.correspondencePath || 'Папка хранения ещё не выбрана'}</span><button onClick={chooseFolder} title={project.correspondencePath ? 'Сменить папку' : 'Выбрать папку'}><Pencil /></button></div></div><div className="section-actions">{project.correspondencePath && <><button className="quiet-button" onClick={() => void scan()} disabled={loading}><Clock3 /> Обновить</button><button className="primary-button" onClick={() => setAddingDraft(true)}><Plus /> Заготовка письма</button></>}</div></div>
    {error && <div className="correspondence-error">{error}</div>}
    {!project.correspondencePath ? <SectionPlaceholder icon={FolderOpen} title="Выберите папку переписки" text="Планер запомнит её для этого проекта и покажет находящиеся внутри папки писем." /> : loading ? <div className="correspondence-loading">Сканирую папку…</div> : letters.length ? <div className="correspondence-table-wrap"><table className="correspondence-table"><thead><tr><th>Дата</th><th>Направление</th><th>Номер</th><th>Корреспондент</th><th>Тема</th><th /></tr></thead><tbody>{letters.map(letter => <tr key={letter.path} data-letter-path={letter.direction === 'outgoing' ? letter.path : undefined} title={letter.direction === 'outgoing' ? 'Сюда можно перетащить файл исходящего письма' : undefined}><td className={letter.draft ? 'draft-date' : ''}>{letter.date}</td><td><span className={`letter-direction ${letter.direction}`}>{letter.direction === 'incoming' ? 'Входящее' : letter.direction === 'outgoing' ? 'Исходящее' : 'Не определено'}</span></td><td><strong>{letter.number}</strong></td><td>{letter.correspondent}</td><td><span className="letter-subject">{letter.subject}</span>{letter.draft && <small className="draft-badge">Заготовка</small>}</td><td><div className="letter-row-actions"><button className="letter-open-button" onClick={() => openFileLink(letter.path)} title="Открыть папку"><FolderOpen /></button><div className="table-actions"><button title="Редактировать" onClick={() => setEditingLetter(letter)}><Pencil /></button><button className="danger" title="Удалить" onClick={() => void deleteLetter(letter)}><Trash2 /></button></div></div></td></tr>)}</tbody></table></div> : <SectionPlaceholder icon={MessageSquareText} title="Писем пока нет" text="В выбранной папке нет вложенных папок писем. Создайте первую заготовку." />}
    {addingDraft && <CorrespondenceDraftEditor root={project.correspondencePath} close={() => setAddingDraft(false)} saved={() => { setAddingDraft(false); void scan() }} />}
    {editingLetter && <CorrespondenceEditor letter={editingLetter} close={() => setEditingLetter(null)} saved={updated => { setFolders(current => current.map(folder => folder.path === editingLetter.path ? updated : folder).sort((left, right) => right.name.localeCompare(left.name))); setEditingLetter(null) }} />}
  </div>
}

function CorrespondenceEditor({ letter, close, saved }: { letter: ParsedCorrespondence; close: () => void; saved: (folder: CorrespondenceFolderRecord) => void }) {
  const displayedDate = letter.date.split('.')
  const initialDate = displayedDate.length === 3 ? `${displayedDate[2]}.${displayedDate[1]}.${displayedDate[0]}` : ''
  const [date, setDate] = useState(initialDate)
  const [number, setNumber] = useState(letter.number === '—' ? '' : letter.number)
  const [direction, setDirection] = useState<'incoming' | 'outgoing'>(letter.direction === 'incoming' ? 'incoming' : 'outgoing')
  const [correspondent, setCorrespondent] = useState(letter.correspondent === 'Формат не распознан' ? '' : letter.correspondent)
  const [subject, setSubject] = useState(letter.subject)
  const [error, setError] = useState('')
  const submit = async (event: React.FormEvent) => {
    event.preventDefault()
    setError('')
    try {
      const updated = await invoke<CorrespondenceFolderRecord>('rename_correspondence_folder', { folderPath: letter.path, date, number, direction, correspondent, subject })
      saved(updated)
    } catch (renameError) {
      setError(String(renameError))
    }
  }
  const route = direction === 'outgoing' ? `ЛАБС - ${correspondent || 'АА'}` : `${correspondent || 'АА'} - ЛАБС`
  return <div className="modal-backdrop" onMouseDown={event => event.target === event.currentTarget && close()}><form className="modal correspondence-editor" onSubmit={submit}><div className="modal-head"><div><span>Папка письма</span><h2>Редактирование</h2></div><button type="button" onClick={close}><X /></button></div><div className="correspondence-form"><label>Дата<input value={date} onChange={event => setDate(event.target.value)} placeholder="2026.08.28 или 2026.08.xx" required /></label><label>Номер<input value={number} onChange={event => setNumber(event.target.value)} placeholder="СО-2026-442" required /></label><label>Направление<select value={direction} onChange={event => setDirection(event.target.value as 'incoming' | 'outgoing')}><option value="outgoing">Исходящее</option><option value="incoming">Входящее</option></select></label><label>Корреспондент<input value={correspondent} onChange={event => setCorrespondent(event.target.value)} placeholder="Например, АА" required /></label><label className="span-2">Тема<input value={subject} onChange={event => setSubject(event.target.value)} required /></label></div><p className="draft-preview">[{date || '2026.08.xx'}] [{number || 'СО-2026-xxx'}] {route}. {subject || 'Тема письма'}</p>{error && <div className="correspondence-error">{error}</div>}<div className="modal-actions"><button type="button" className="quiet-button" onClick={close}>Отмена</button><button type="submit" className="primary-button">Сохранить</button></div></form></div>
}

function CorrespondenceDraftEditor({ root, close, saved }: { root: string; close: () => void; saved: () => void }) {
  const now = new Date()
  const [year, setYear] = useState(now.getFullYear())
  const [month, setMonth] = useState(now.getMonth() + 1)
  const [counterparty, setCounterparty] = useState('')
  const [subject, setSubject] = useState('')
  const [error, setError] = useState('')
  const submit = async (event: React.FormEvent) => {
    event.preventDefault()
    setError('')
    try {
      await invoke('create_correspondence_draft', { root, year, month, counterparty, subject })
      saved()
    } catch (createError) {
      setError(String(createError))
    }
  }
  return <div className="modal-backdrop" onMouseDown={event => event.target === event.currentTarget && close()}><form className="modal correspondence-editor" onSubmit={submit}><div className="modal-head"><div><span>Исходящее письмо</span><h2>Новая заготовка</h2></div><button type="button" onClick={close}><X /></button></div><div className="correspondence-form"><label>Год<input type="number" min="2000" max="2100" value={year} onChange={event => setYear(Number(event.target.value))} required /></label><label>Месяц<input type="number" min="1" max="12" value={month} onChange={event => setMonth(Number(event.target.value))} required /></label><label className="span-2">Адресат<input value={counterparty} onChange={event => setCounterparty(event.target.value)} placeholder="Например, АА" required /></label><label className="span-2">Тема письма<input autoFocus value={subject} onChange={event => setSubject(event.target.value)} placeholder="Например, По поводу этапа 1.2" required /></label></div><p className="draft-preview">[{year}.{String(month).padStart(2, '0')}.xx] [СО-{year}-xxx] ЛАБС - {counterparty || 'АА'}. {subject || 'Тема письма'}</p>{error && <div className="correspondence-error">{error}</div>}<div className="modal-actions"><button type="button" className="quiet-button" onClick={close}>Отмена</button><button type="submit" className="primary-button">Создать папку</button></div></form></div>
}

function ProjectTasks({ project, data, setData }: { project: Project; data: PlannerData; setData: React.Dispatch<React.SetStateAction<PlannerData>> }) {
  const [editing, setEditing] = useState<Task | null>(null)
  const [adding, setAdding] = useState(false)
  const [showCompleted, setShowCompleted] = useState(false)
  const tasks = data.tasks.filter(task => task.projectId === project.id)
  const completedCount = tasks.filter(task => task.done).length
  const visibleProjectTasks = showCompleted ? tasks : tasks.filter(task => !task.done)
  const toggle = (id: string) => setData(current => ({ ...current, tasks: current.tasks.map(task => task.id === id ? { ...task, done: !task.done, status: !task.done ? 'done' : 'planned' } : task) }))
  return <div className="workspace-section"><div className="section-heading"><div><h2>Задачи</h2><p>{tasks.filter(task => !task.done).length} открытых · {completedCount} завершённых</p></div><div className="section-actions">{completedCount > 0 && <button className="quiet-button completed-toggle" onClick={() => setShowCompleted(value => !value)}>{showCompleted ? <EyeOff /> : <Eye />}{showCompleted ? 'Скрыть выполненные' : `Показать выполненные (${completedCount})`}</button>}<button className="primary-button" onClick={() => setAdding(true)}><Plus /> Добавить задачу</button></div></div>{visibleProjectTasks.length ? <div className="project-task-cards">{visibleProjectTasks.map(task => <article key={task.id} className={task.done ? 'done' : ''}><button className="check-button" onClick={() => toggle(task.id)}>{task.done ? <CheckCircle2 /> : <Circle />}</button><div className="project-task-copy"><div><strong>{task.title}</strong><span className={`task-status ${task.status}`}>{taskStatusLabel[task.status]}</span></div>{task.description && <p>{task.description}</p>}<footer><span><CalendarDays /> {task.due ? formatContractDate(task.due) : 'Без срока'}</span><span>{task.assignee || 'Исполнитель не указан'}</span><span>{taskSourceLabel[task.source]}</span>{task.contractStageId && <span>Этап {data.contractStages.find(stage => stage.id === task.contractStageId)?.number || '—'}</span>}</footer></div><div className="table-actions visible"><button title="Редактировать" onClick={() => setEditing(task)}><Pencil /></button><button className="danger" title="Удалить" onClick={() => { if (window.confirm(`Удалить задачу «${task.title}»?`)) setData(current => ({ ...current, tasks: current.tasks.filter(item => item.id !== task.id) })) }}><Trash2 /></button></div></article>)}</div> : <SectionPlaceholder icon={ListTodo} title={tasks.length ? 'Все задачи выполнены' : 'Задач пока нет'} text={tasks.length ? 'Включите отображение выполненных задач или создайте новую.' : 'Создайте первую задачу для текущего проекта.'} />}{adding && <TaskEditor project={project} stages={data.contractStages} close={() => setAdding(false)} save={task => { setData(current => ({ ...current, tasks: [task, ...current.tasks] })); setAdding(false) }} />}{editing && <TaskEditor project={project} stages={data.contractStages} task={editing} close={() => setEditing(null)} save={task => { setData(current => ({ ...current, tasks: current.tasks.map(item => item.id === task.id ? task : item) })); setEditing(null) }} />}</div>
}

const taskStatusLabel: Record<TaskStatus, string> = { planned: 'Запланирована', in_progress: 'В работе', blocked: 'Заблокирована', done: 'Завершена' }
const taskSourceLabel: Record<TaskSource, string> = { manual: 'Вручную', meeting: 'Из совещания', outlook: 'Outlook', youtrack: 'YouTrack' }

function TaskEditor({ project, stages, task, save, close }: { project: Project; stages: ContractStage[]; task?: Task; save: (task: Task) => void; close: () => void }) {
  const [form, setForm] = useState<Task>(task || { id: uid('t'), projectId: project.id, title: '', description: '', due: '', priority: 'medium', done: false, status: 'planned', assignee: '', source: 'manual', contractStageId: '', meetingId: '' })
  const update = <K extends keyof Task>(key: K, value: Task[K]) => setForm(current => ({ ...current, [key]: value }))
  const submit = (event: React.FormEvent) => { event.preventDefault(); if (!form.title.trim()) return; const done = form.status === 'done'; save({ ...form, title: form.title.trim(), description: form.description.trim(), assignee: form.assignee.trim(), done }) }
  return <div className="modal-backdrop" onMouseDown={event => event.target === event.currentTarget && close()}><form className="modal task-editor" onSubmit={submit}><div className="modal-head"><div><span>Задача проекта</span><h2>{task ? 'Редактирование задачи' : 'Новая задача'}</h2></div><button type="button" onClick={close}><X /></button></div><div className="task-form"><label className="span-2">Название<input autoFocus value={form.title} onChange={event => update('title', event.target.value)} required /></label><label className="span-2">Описание<textarea value={form.description} onChange={event => update('description', event.target.value)} /></label><label>Срок<input type="date" value={form.due} onChange={event => update('due', event.target.value)} /></label><label>Ответственный<input value={form.assignee} onChange={event => update('assignee', event.target.value)} /></label><label>Статус<select value={form.status} onChange={event => update('status', event.target.value as TaskStatus)}><option value="planned">Запланирована</option><option value="in_progress">В работе</option><option value="blocked">Заблокирована</option><option value="done">Завершена</option></select></label><label>Приоритет<select value={form.priority} onChange={event => update('priority', event.target.value as Priority)}><option value="high">Высокий</option><option value="medium">Обычный</option><option value="low">Низкий</option></select></label><label>Источник<select value={form.source} onChange={event => update('source', event.target.value as TaskSource)}><option value="manual">Вручную</option><option value="meeting">Совещание</option><option value="outlook">Outlook</option><option value="youtrack">YouTrack</option></select></label><label>Этап договора<select value={form.contractStageId} onChange={event => update('contractStageId', event.target.value)}><option value="">Не связан</option>{stages.filter(stage => stage.projectId === project.id).map(stage => <option key={stage.id} value={stage.id}>{stage.number} · {stage.name}</option>)}</select></label></div><div className="modal-actions"><button type="button" className="quiet-button" onClick={close}>Отмена</button><button type="submit" className="primary-button">Сохранить</button></div></form></div>
}

const trackingStatusLabel: Record<TrackingStatus, string> = {
  waiting: 'Ожидаю',
  reminded: 'Напомнил',
  completed: 'Получено',
  cancelled: 'Отменено',
}

function ProjectTrackings({ project, data, setData }: { project: Project; data: PlannerData; setData: React.Dispatch<React.SetStateAction<PlannerData>> }) {
  const [adding, setAdding] = useState(false)
  const [editing, setEditing] = useState<Tracking | null>(null)
  const [showClosed, setShowClosed] = useState(false)
  const today = new Date().toISOString().slice(0, 10)
  const projectTrackings = data.trackings.filter(tracking => tracking.projectId === project.id)
  const isClosed = (tracking: Tracking) => tracking.status === 'completed' || tracking.status === 'cancelled'
  const closedCount = projectTrackings.filter(isClosed).length
  const visible = projectTrackings
    .filter(tracking => showClosed || !isClosed(tracking))
    .sort((a, b) => Number(isClosed(a)) - Number(isClosed(b)) || (a.checkDate || '9999').localeCompare(b.checkDate || '9999'))
  const remove = (tracking: Tracking) => {
    if (window.confirm(`Удалить отслеживание «${tracking.subject}»?`)) {
      setData(current => ({ ...current, trackings: current.trackings.filter(item => item.id !== tracking.id) }))
    }
  }
  const complete = (id: string) => setData(current => ({ ...current, trackings: current.trackings.map(item => item.id === id ? { ...item, status: 'completed' } : item) }))
  return <div className="workspace-section trackings-section">
    <div className="section-heading"><div><h2>Отслеживания</h2><p>Запросы и поручения другим людям, по которым вы ждёте результат.</p></div><div className="section-actions">{closedCount > 0 && <button className="quiet-button completed-toggle" onClick={() => setShowClosed(value => !value)}>{showClosed ? <EyeOff /> : <Eye />}{showClosed ? 'Скрыть закрытые' : `Показать закрытые (${closedCount})`}</button>}<button className="primary-button" onClick={() => setAdding(true)}><Plus /> Добавить отслеживание</button></div></div>
    {visible.length ? <div className="tracking-list">{visible.map(tracking => {
      const overdue = !isClosed(tracking) && Boolean(tracking.checkDate) && tracking.checkDate < today
      return <article key={tracking.id} className={`${overdue ? 'overdue' : ''} ${isClosed(tracking) ? 'closed' : ''}`}>
        <div className="tracking-copy"><div className="tracking-head"><strong>{tracking.subject}</strong><span className={`tracking-status ${tracking.status}`}>{trackingStatusLabel[tracking.status]}</span></div>
          <div className="tracking-meta"><span>Кому: <b>{tracking.recipient || 'не указано'}</b></span><span>Отправлено: <b>{formatContractDate(tracking.sentDate)}</b></span><span className={overdue ? 'overdue' : ''}>{overdue ? 'Просрочена проверка: ' : 'Проверить: '}<b>{formatContractDate(tracking.checkDate)}</b></span>{tracking.sourceEmailId && <span>Outlook</span>}</div>
          {tracking.comment && <p>{tracking.comment}</p>}
        </div>
        <div className="tracking-controls">{!isClosed(tracking) && <button className="tracking-complete" onClick={() => complete(tracking.id)}><Check /> Получено</button>}<div className="table-actions visible"><button title="Редактировать" onClick={() => setEditing(tracking)}><Pencil /></button><button className="danger" title="Удалить" onClick={() => remove(tracking)}><Trash2 /></button></div></div>
      </article>
    })}</div> : <SectionPlaceholder icon={BellRing} title={projectTrackings.length ? 'Все отслеживания закрыты' : 'Отслеживаний пока нет'} text={projectTrackings.length ? 'Включите отображение закрытых записей или создайте новую.' : 'Добавьте запрос, по которому нужно получить результат или напомнить адресату.'} />}
    {adding && <TrackingEditor projectId={project.id} close={() => setAdding(false)} save={tracking => { setData(current => ({ ...current, trackings: [tracking, ...current.trackings] })); setAdding(false) }} />}
    {editing && <TrackingEditor projectId={project.id} tracking={editing} close={() => setEditing(null)} save={tracking => { setData(current => ({ ...current, trackings: current.trackings.map(item => item.id === tracking.id ? tracking : item) })); setEditing(null) }} />}
  </div>
}

function TrackingEditor({ projectId, tracking, save, close }: { projectId: string; tracking?: Tracking; save: (tracking: Tracking) => void; close: () => void }) {
  const today = new Date().toISOString().slice(0, 10)
  const defaultCheck = new Date()
  defaultCheck.setDate(defaultCheck.getDate() + 7)
  const [form, setForm] = useState<Tracking>(tracking || { id: uid('tr'), projectId, subject: '', recipient: '', sentDate: today, checkDate: defaultCheck.toISOString().slice(0, 10), status: 'waiting', comment: '', sourceEmailId: '' })
  const update = <K extends keyof Tracking>(key: K, value: Tracking[K]) => setForm(current => ({ ...current, [key]: value }))
  const submit = (event: React.FormEvent) => {
    event.preventDefault()
    if (!form.subject.trim()) return
    save({ ...form, subject: form.subject.trim(), recipient: form.recipient.trim(), comment: form.comment.trim() })
  }
  return <div className="modal-backdrop" onMouseDown={event => event.target === event.currentTarget && close()}><form className="modal tracking-editor" onSubmit={submit}><div className="modal-head"><div><span>Запрос другому человеку</span><h2>{tracking ? 'Редактирование отслеживания' : 'Новое отслеживание'}</h2></div><button type="button" onClick={close}><X /></button></div><div className="tracking-form"><label className="span-2">Что ожидаем<input autoFocus value={form.subject} onChange={event => update('subject', event.target.value)} required /></label><label className="span-2">Адресат<input value={form.recipient} onChange={event => update('recipient', event.target.value)} placeholder="Имя, компания или адрес электронной почты" /></label><label>Дата отправки<input type="date" value={form.sentDate} onChange={event => update('sentDate', event.target.value)} /></label><label>Дата следующей проверки<input type="date" value={form.checkDate} onChange={event => update('checkDate', event.target.value)} /></label><label>Статус<select value={form.status} onChange={event => update('status', event.target.value as TrackingStatus)}><option value="waiting">Ожидаю</option><option value="reminded">Напомнил</option><option value="completed">Получено</option><option value="cancelled">Отменено</option></select></label><label className="span-2">Комментарий<textarea value={form.comment} onChange={event => update('comment', event.target.value)} placeholder="Что было запрошено и что нужно проверить" /></label></div><div className="modal-actions"><button type="button" className="quiet-button" onClick={close}>Отмена</button><button type="submit" className="primary-button">Сохранить</button></div></form></div>
}

function ProjectMeetings({ project, data, setData }: { project: Project; data: PlannerData; setData: React.Dispatch<React.SetStateAction<PlannerData>> }) {
  const [adding, setAdding] = useState(false)
  const [editing, setEditing] = useState<Meeting | null>(null)
  const meetings = data.meetings.filter(meeting => meeting.projectId === project.id).sort((a, b) => `${b.date}${b.time}`.localeCompare(`${a.date}${a.time}`))
  const createTask = (meeting: Meeting, action: MeetingAction) => {
    const task: Task = { id: uid('t'), projectId: project.id, title: action.title, description: `Поручение из совещания «${meeting.topic}»`, due: action.due, priority: 'medium', done: false, status: 'planned', assignee: action.assignee, source: 'meeting', contractStageId: '', meetingId: meeting.id }
    setData(current => ({ ...current, tasks: [task, ...current.tasks], meetings: current.meetings.map(item => item.id === meeting.id ? { ...item, actions: item.actions.map(entry => entry.id === action.id ? { ...entry, taskId: task.id } : entry) } : item) }))
  }
  return <div className="workspace-section meetings-section"><div className="section-heading"><div><span>Протоколы и решения</span><h2>Совещания</h2><p>{meetings.length} записей по проекту</p></div><button className="primary-button" onClick={() => setAdding(true)}><Plus /> Новое совещание</button></div>{meetings.length ? <div className="meeting-list">{meetings.map(meeting => <article key={meeting.id}><header><div><time>{formatContractDate(meeting.date)}{meeting.time && ` · ${meeting.time}`}</time><h3>{meeting.topic}</h3><span>{meeting.format === 'online' ? 'Онлайн' : meeting.format === 'offline' ? 'Очно' : 'Гибридно'}{meeting.participants && ` · ${meeting.participants}`}</span></div><div className="table-actions visible"><button onClick={() => setEditing(meeting)}><Pencil /></button><button className="danger" onClick={() => { if (window.confirm(`Удалить совещание «${meeting.topic}»?`)) setData(current => ({ ...current, meetings: current.meetings.filter(item => item.id !== meeting.id) })) }}><Trash2 /></button></div></header>{meeting.agenda && <MeetingText title="Повестка" text={meeting.agenda} />}{meeting.protocol && <MeetingText title="Протокол" text={meeting.protocol} />}{meeting.decisions && <MeetingText title="Решения" text={meeting.decisions} />}{meeting.actions.length > 0 && <div className="meeting-actions"><h4>Поручения</h4>{meeting.actions.map(action => <div key={action.id}><div><strong>{action.title}</strong><span>{action.assignee || 'Исполнитель не указан'}{action.due && ` · до ${formatContractDate(action.due)}`}</span></div>{action.taskId ? <span className="task-created"><Check /> Задача создана</span> : <button onClick={() => createTask(meeting, action)}><Plus /> Создать задачу</button>}</div>)}</div>}</article>)}</div> : <SectionPlaceholder icon={CalendarDays} title="Совещаний пока нет" text="Создайте запись встречи, добавьте протокол, решения и поручения." />}{adding && <MeetingEditor projectId={project.id} close={() => setAdding(false)} save={meeting => { setData(current => ({ ...current, meetings: [meeting, ...current.meetings] })); setAdding(false) }} />}{editing && <MeetingEditor projectId={project.id} meeting={editing} close={() => setEditing(null)} save={meeting => { setData(current => ({ ...current, meetings: current.meetings.map(item => item.id === meeting.id ? meeting : item) })); setEditing(null) }} />}</div>
}

function MeetingText({ title, text }: { title: string; text: string }) { return <section className="meeting-text"><h4>{title}</h4><p>{text}</p></section> }

function MeetingEditor({ projectId, meeting, save, close }: { projectId: string; meeting?: Meeting; save: (meeting: Meeting) => void; close: () => void }) {
  const [form, setForm] = useState<Meeting>(meeting ? { ...meeting, actions: meeting.actions.map(action => ({ ...action })) } : { id: uid('m'), projectId, date: new Date().toISOString().slice(0, 10), time: '', format: 'online', topic: '', participants: '', agenda: '', protocol: '', decisions: '', actions: [] })
  const update = <K extends keyof Meeting>(key: K, value: Meeting[K]) => setForm(current => ({ ...current, [key]: value }))
  const updateAction = (id: string, patch: Partial<MeetingAction>) => update('actions', form.actions.map(action => action.id === id ? { ...action, ...patch } : action))
  const submit = (event: React.FormEvent) => { event.preventDefault(); if (!form.topic.trim()) return; save({ ...form, topic: form.topic.trim(), participants: form.participants.trim(), agenda: form.agenda.trim(), protocol: form.protocol.trim(), decisions: form.decisions.trim(), actions: form.actions.filter(action => action.title.trim()).map(action => ({ ...action, title: action.title.trim(), assignee: action.assignee.trim() })) }) }
  return <div className="modal-backdrop" onMouseDown={event => event.target === event.currentTarget && close()}><form className="modal meeting-editor" onSubmit={submit}><div className="modal-head"><div><span>Совещание проекта</span><h2>{meeting ? 'Редактирование протокола' : 'Новое совещание'}</h2></div><button type="button" onClick={close}><X /></button></div><div className="meeting-form"><label className="span-3">Тема<input autoFocus value={form.topic} onChange={event => update('topic', event.target.value)} required /></label><label>Дата<input type="date" value={form.date} onChange={event => update('date', event.target.value)} /></label><label>Время<input type="time" value={form.time} onChange={event => update('time', event.target.value)} /></label><label>Формат<select value={form.format} onChange={event => update('format', event.target.value as Meeting['format'])}><option value="online">Онлайн</option><option value="offline">Очно</option><option value="hybrid">Гибридно</option></select></label><label className="span-3">Участники<input value={form.participants} onChange={event => update('participants', event.target.value)} placeholder="Имена или организации через запятую" /></label><label className="span-3">Повестка<textarea value={form.agenda} onChange={event => update('agenda', event.target.value)} /></label><label className="span-3">Протокол<textarea value={form.protocol} onChange={event => update('protocol', event.target.value)} /></label><label className="span-3">Принятые решения<textarea value={form.decisions} onChange={event => update('decisions', event.target.value)} /></label><div className="action-editor span-3"><div><strong>Поручения</strong><button type="button" onClick={() => update('actions', [...form.actions, { id: uid('ma'), title: '', assignee: '', due: '', taskId: '' }])}><Plus /> Добавить поручение</button></div>{form.actions.map(action => <div className="action-editor-row" key={action.id}><input value={action.title} onChange={event => updateAction(action.id, { title: event.target.value })} placeholder="Что нужно сделать" /><input value={action.assignee} onChange={event => updateAction(action.id, { assignee: event.target.value })} placeholder="Исполнитель" /><input type="date" value={action.due} onChange={event => updateAction(action.id, { due: event.target.value })} /><button type="button" onClick={() => update('actions', form.actions.filter(item => item.id !== action.id))}><X /></button></div>)}</div></div><div className="modal-actions"><button type="button" className="quiet-button" onClick={close}>Отмена</button><button type="submit" className="primary-button">Сохранить</button></div></form></div>
}

const contractStatusLabel: Record<ContractStageStatus, string> = { not_started: 'Не начат', started: 'Начат', suspended: 'Приостановлен', work_completed: 'Работы выполнены', act_sent: 'Акт отправлен', act_signed: 'Акт подписан', closed: 'Закрыт' }

function ContractTable({ stages, edit, remove }: { stages: ContractStage[]; edit: (stage: ContractStage) => void; remove: (stage: ContractStage) => void }) {
  if (!stages.length) return <SectionPlaceholder icon={FileText} title="Этапов пока нет" text="Добавьте первый этап договора, чтобы сформировать график выполнения работ." />
  return <div className="contract-table-wrap"><table className="contract-table"><thead><tr><th>№ этапа</th><th>Название этапа</th><th>Состав работ</th><th>Начало</th><th>Завершение</th><th>Статус</th><th>Оплата</th><th>Комментарий</th><th aria-label="Действия" /></tr></thead><tbody>{stages.map(stage => <tr key={stage.id}><td><b>{stage.number}</b></td><td><strong>{stage.name}</strong></td><td className="scope-cell">{stage.scope || '—'}</td><td className="date-cell">{formatContractDate(stage.startDate)}</td><td className="date-cell">{formatContractDate(stage.endDate)}</td><td><span className={`contract-status ${stage.status} ${stage.status === 'suspended' ? 'has-tooltip' : ''}`} data-tooltip={stage.status === 'suspended' ? (stage.suspensionLetter || 'Письмо не указано') : undefined}>{contractStatusLabel[stage.status]}</span></td><td><PaymentSummary stage={stage} /></td><td className="comment-cell">{stage.comment || '—'}</td><td><div className="table-actions"><button title="Редактировать" onClick={() => edit(stage)}><Pencil /></button><button className="danger" title="Удалить" onClick={() => remove(stage)}><Trash2 /></button></div></td></tr>)}</tbody></table></div>
}

function PaymentSummary({ stage }: { stage: ContractStage }) {
  const advanceDanger = stage.advanceStatus === 'unpaid' && stage.status !== 'not_started'
  const finalDanger = stage.finalPaymentStatus === 'unpaid' && (stage.status === 'act_signed' || stage.status === 'closed')
  const advanceValue = (Number(stage.cost) || 0) * (Number(stage.advancePercent) || 0) / 100
  const finalValue = Math.max(0, (Number(stage.cost) || 0) - advanceValue)
  return <div className="payment-summary"><span><span>Аванс</span><em>{stage.advancePercent || 0}% · {formatMoney(advanceValue)}</em><b className={stage.advanceStatus === 'paid' ? 'paid' : advanceDanger ? 'unpaid-danger' : 'unpaid-neutral'}>{stage.advanceStatus === 'paid' ? 'Оплачен' : 'Не оплачен'}</b></span><span><span>ОКР</span><em>{formatMoney(finalValue)}</em><b className={stage.finalPaymentStatus === 'paid' ? 'paid' : finalDanger ? 'unpaid-danger' : 'unpaid-neutral'}>{stage.finalPaymentStatus === 'paid' ? 'Оплачен' : 'Не оплачен'}</b></span></div>
}

function formatMoney(value: number | undefined) {
  return new Intl.NumberFormat('ru-RU', { style: 'currency', currency: 'RUB', maximumFractionDigits: 2 }).format(Number(value) || 0)
}

function formatContractDate(value: string) {
  return value ? new Date(`${value}T00:00:00`).toLocaleDateString('ru-RU') : '—'
}

function ContractStageEditor({ projectId, stage, save, close }: { projectId: string; stage?: ContractStage; save: (stage: ContractStage) => void; close: () => void }) {
  const legacyStatusMigration: Record<string, ContractStageStatus> = { planned: 'not_started', in_progress: 'started', completed: 'work_completed', paused: 'suspended' }
  const [form, setForm] = useState<ContractStage>(() => ({
    id: stage?.id || uid('cs'),
    projectId,
    number: stage?.number || '',
    name: stage?.name || '',
    scope: stage?.scope || '',
    startDate: stage?.startDate || '',
    endDate: stage?.endDate || '',
    status: legacyStatusMigration[stage?.status || ''] || stage?.status || 'not_started',
    suspensionLetter: stage?.suspensionLetter || '',
    cost: Number(stage?.cost) || 0,
    advancePercent: Number(stage?.advancePercent) || 0,
    advanceStatus: stage?.advanceStatus || 'unpaid',
    finalPaymentStatus: stage?.finalPaymentStatus || 'unpaid',
    comment: stage?.comment || '',
  }))
  const update = <K extends keyof ContractStage>(key: K, value: ContractStage[K]) => setForm(current => ({ ...current, [key]: value }))
  const submit = (event: React.FormEvent) => { event.preventDefault(); if (!form.number.trim() || !form.name.trim()) return; save({ ...form, number: form.number.trim(), name: form.name.trim(), scope: (form.scope || '').trim(), suspensionLetter: (form.suspensionLetter || '').trim(), comment: (form.comment || '').trim() }) }
  return <div className="modal-backdrop" onMouseDown={event => event.target === event.currentTarget && close()}><form className="modal contract-editor" onSubmit={submit}><div className="modal-head"><div><span>Договор</span><h2>{stage ? 'Редактирование этапа' : 'Новый этап'}</h2></div><button type="button" onClick={close}><X /></button></div><div className="contract-form"><label>Номер этапа<input autoFocus value={form.number} onChange={event => update('number', event.target.value)} required /></label><label className="span-2">Название этапа<input value={form.name} onChange={event => update('name', event.target.value)} required /></label><label className="span-3">Состав работ<textarea value={form.scope} onChange={event => update('scope', event.target.value)} /></label><label>Срок начала<input type="date" value={form.startDate} onChange={event => update('startDate', event.target.value)} /></label><label>Срок завершения<input type="date" value={form.endDate} onChange={event => update('endDate', event.target.value)} /></label><label>Статус<select value={form.status} onChange={event => update('status', event.target.value as ContractStageStatus)}><option value="not_started">Не начат</option><option value="started">Начат</option><option value="suspended">Приостановлен</option><option value="work_completed">Работы выполнены</option><option value="act_sent">Акт отправлен</option><option value="act_signed">Акт подписан</option><option value="closed">Закрыт</option></select></label><label>Стоимость<input type="number" min="0" step="0.01" value={form.cost} onChange={event => update('cost', Number(event.target.value))} /></label><label>Размер аванса, %<input type="number" min="0" max="100" step="0.01" value={form.advancePercent} onChange={event => update('advancePercent', Math.min(100, Math.max(0, Number(event.target.value))))} /></label>{form.status === 'suspended' && <label className="span-3 suspension-field">Номер письма<input value={form.suspensionLetter} onChange={event => update('suspensionLetter', event.target.value)} placeholder="Например, исх. № 125 от 14.08.2026" /><small>Будет показан в подсказке при наведении на статус этапа.</small></label>}<div className="payment-fields span-3"><div><span>Аванс</span><PaymentSelect value={form.advanceStatus} onChange={value => update('advanceStatus', value)} /></div><div><span>ОКР</span><PaymentSelect value={form.finalPaymentStatus} onChange={value => update('finalPaymentStatus', value)} /></div></div><label className="span-3">Комментарий<textarea value={form.comment} onChange={event => update('comment', event.target.value)} /></label></div><div className="modal-actions"><button type="button" className="quiet-button" onClick={close}>Отмена</button><button type="submit" className="primary-button">{stage ? 'Сохранить' : 'Добавить этап'}</button></div></form></div>
}

function PaymentSelect({ value, onChange }: { value: PaymentStatus; onChange: (value: PaymentStatus) => void }) {
  return <select value={value} onChange={event => onChange(event.target.value as PaymentStatus)}><option value="unpaid">Не оплачен</option><option value="paid">Оплачен</option></select>
}

function reorderProjectLinks(data: PlannerData, projectId: string, sourceId: string, targetId: string): PlannerData {
  const ordered = data.links.filter(link => link.projectId === projectId)
  const sourceIndex = ordered.findIndex(link => link.id === sourceId)
  const targetIndex = ordered.findIndex(link => link.id === targetId)
  if (sourceIndex < 0 || targetIndex < 0) return data
  const [moved] = ordered.splice(sourceIndex, 1)
  ordered.splice(targetIndex, 0, moved)
  let projectLinkIndex = 0
  return { ...data, links: data.links.map(link => link.projectId === projectId ? ordered[projectLinkIndex++] : link) }
}

function safeWebUrl(target: string) {
  const value = target.trim()
  return /^https?:\/\//i.test(value) ? value : `https://${value}`
}

function LinkEditor({ projectId, link, save, close }: { projectId: string; link?: PlannerData['links'][number]; save: (link: PlannerData['links'][number]) => void; close: () => void }) {
  const [kind, setKind] = useState<'web' | 'file'>(link?.kind || 'web')
  const [label, setLabel] = useState(link?.label || '')
  const [target, setTarget] = useState(link?.target || '')
  const submit = (event: React.FormEvent) => {
    event.preventDefault()
    if (!label.trim() || !target.trim()) return
    save({ id: link?.id || uid('l'), projectId, label: label.trim(), target: kind === 'web' ? safeWebUrl(target) : target.trim(), kind })
  }
  return <div className="modal-backdrop" onMouseDown={event => event.target === event.currentTarget && close()}>
    <form className="modal link-editor" onSubmit={submit}>
      <div className="modal-head"><div><span>Материалы проекта</span><h2>{link ? 'Редактирование ссылки' : 'Новая ссылка'}</h2></div><button type="button" onClick={close}><X /></button></div>
      <div className="type-switch"><button type="button" className={kind === 'web' ? 'active' : ''} onClick={() => setKind('web')}><ExternalLink /> Веб-ссылка</button><button type="button" className={kind === 'file' ? 'active' : ''} onClick={() => setKind('file')}><FolderOpen /> Файл или папка</button></div>
      <div className="link-form">
        <label>Название<input autoFocus value={label} onChange={event => setLabel(event.target.value)} placeholder={kind === 'web' ? 'Например, YouTrack' : 'Например, Папка проекта'} required /></label>
        <label>{kind === 'web' ? 'Адрес ссылки' : 'Путь к файлу или папке'}<input value={target} onChange={event => setTarget(event.target.value)} placeholder={kind === 'web' ? 'https://example.com' : 'C:\\Projects\\Project-01'} required /></label>
        {kind === 'file' && <p className="form-hint"><FolderOpen /> В настольной версии файл или папка откроются системным приложением по умолчанию.</p>}
      </div>
      <div className="modal-actions"><button type="button" className="quiet-button" onClick={close}>Отмена</button><button type="submit" className="primary-button">{link ? 'Сохранить' : 'Добавить'}</button></div>
    </form>
  </div>
}

function SectionPlaceholder({ icon: Icon, title, text }: { icon: typeof Link2; title: string; text: string }) {
  return <div className="section-placeholder"><div><Icon /></div><h2>{title}</h2><p>{text}</p><span>Раздел будет дополнен позже</span></div>
}

function Tasks({ data, tasks, setData, toggleTask }: { data: PlannerData; tasks: PlannerData['tasks']; setData: React.Dispatch<React.SetStateAction<PlannerData>>; toggleTask: (id: string) => void }) {
  const add = () => setData(d => ({ ...d, tasks: [{ id: uid('t'), title: 'Новая задача', projectId: d.projects[0]?.id || '', due: new Date().toISOString().slice(0,10), priority: 'medium', done: false, description: '', status: 'planned', assignee: '', source: 'manual', contractStageId: '', meetingId: '' }, ...d.tasks] }))
  return <><PageIntro eyebrow="Рабочий список" title="Задачи" subtitle={`${data.tasks.filter(t => !t.done).length} открытых задач во всех проектах.`} action={<button className="primary-button" onClick={add}><Plus /> Новая задача</button>} /><div className="panel standalone-list"><div className="list-tabs"><button className="active">Все</button><button>Сегодня</button><button>Предстоящие</button><button>Завершённые</button></div>{tasks.map(task => <TaskRow key={task.id} task={task} project={data.projects.find(p => p.id === task.projectId)} toggle={() => toggleTask(task.id)} />)}</div></>
}

function Notes({ data, setData }: { data: PlannerData; setData: React.Dispatch<React.SetStateAction<PlannerData>> }) {
  const [text, setText] = useState('')
  const add = () => { if (!text.trim()) return; setData(d => ({ ...d, notes: [{ id: uid('n'), text: text.trim(), createdAt: new Date().toISOString() }, ...d.notes] })); setText('') }
  return <><PageIntro eyebrow="Блокнот" title="Заметки" subtitle="Идеи, наблюдения и важные детали." /><div className="note-composer"><textarea value={text} onChange={e => setText(e.target.value)} placeholder="Запишите мысль…" /><div><span><MessageSquareText /> Можно превратить в задачу позже</span><button className="primary-button" onClick={add}>Сохранить</button></div></div><div className="notes-grid">{data.notes.map(n => <article key={n.id}><StickyNote /><p>{n.text}</p><footer><span>{data.projects.find(p => p.id === n.projectId)?.name || 'Личная заметка'}</span><time>{new Date(n.createdAt).toLocaleDateString('ru-RU')}</time></footer></article>)}</div></>
}

function QuickAdd({ data, setData, close }: { data: PlannerData; setData: React.Dispatch<React.SetStateAction<PlannerData>>; close: () => void }) {
  const [type, setType] = useState<'task'|'note'>('task'); const [text, setText] = useState(''); const [projectId, setProject] = useState(data.projects[0]?.id || '')
  const save = () => { if (!text.trim()) return; setData(d => type === 'task' ? ({ ...d, tasks: [{ id: uid('t'), title: text.trim(), projectId, due: new Date().toISOString().slice(0,10), priority: 'medium', done: false, description: '', status: 'planned', assignee: '', source: 'manual', contractStageId: '', meetingId: '' }, ...d.tasks] }) : ({ ...d, notes: [{ id: uid('n'), text: text.trim(), projectId, createdAt: new Date().toISOString() }, ...d.notes] })); close() }
  return <div className="modal-backdrop" onMouseDown={e => e.target === e.currentTarget && close()}><div className="modal"><div className="modal-head"><div><span>Быстрое добавление</span><h2>Зафиксировать важное</h2></div><button onClick={close}><X /></button></div><div className="type-switch"><button className={type === 'task' ? 'active' : ''} onClick={() => setType('task')}><ListTodo /> Задача</button><button className={type === 'note' ? 'active' : ''} onClick={() => setType('note')}><StickyNote /> Заметка</button></div><textarea autoFocus value={text} onChange={e => setText(e.target.value)} placeholder={type === 'task' ? 'Что нужно сделать?' : 'О чём важно не забыть?'} /><label>Проект<select value={projectId} onChange={e => setProject(e.target.value)}>{data.projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}</select></label><div className="modal-actions"><button className="quiet-button" onClick={close}>Отмена</button><button className="primary-button" onClick={save}>Добавить</button></div></div></div>
}

function CompactView({ data, setData, onExpand, onHide, native }: { data: PlannerData; setData: React.Dispatch<React.SetStateAction<PlannerData>>; onExpand: () => void; onHide?: () => void; native?: boolean }) {
  const [text, setText] = useState(''); const tasks = useMemo(() => data.tasks.filter(t => !t.done).slice(0, 3), [data.tasks])
  const add = () => { if (!text.trim()) return; setData(d => ({ ...d, notes: [{ id: uid('n'), text: text.trim(), createdAt: new Date().toISOString() }, ...d.notes] })); setText('') }
  const startDrag = async (event: React.MouseEvent<HTMLElement>) => { if (native && event.button === 0 && !(event.target as Element).closest('button')) await getCurrentWindow().startDragging() }
  return <div className="compact-page"><div className="compact-card"><header data-tauri-drag-region onMouseDown={startDrag}><div className="brand" data-tauri-drag-region><div className="brand-mark"><Check /></div><span>myplanner</span></div><div className="sticker-window-controls"><button title="Открыть MyPlanner" onClick={onExpand}><ArrowUpRight /></button>{onHide && <button title="Скрыть стикер" onClick={onHide}><X /></button>}</div></header><div className="compact-date"><span>Сегодня</span><strong>{new Date().toLocaleDateString('ru-RU', { day: 'numeric', month: 'long' })}</strong></div><div className="compact-tasks">{tasks.map(t => <button key={t.id} onClick={() => setData(d => ({ ...d, tasks: d.tasks.map(x => x.id === t.id ? { ...x, done: true, status: 'done' } : x) }))}><Circle /><span>{t.title}<small>{data.projects.find(p => p.id === t.projectId)?.name}</small></span></button>)}</div><div className="sticky-input"><textarea value={text} onChange={e => setText(e.target.value)} placeholder="Быстрая заметка…" /><button onClick={add}><Plus /></button></div><footer><span>{tasks.length} ближайшие задачи</span><button onClick={onExpand}>Открыть планнер <ChevronRight /></button></footer></div></div>
}

export default App
