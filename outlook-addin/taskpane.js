const bridgeUrl = 'http://127.0.0.1:17832/capture'
const keyInput = document.querySelector('#key')
const button = document.querySelector('#capture')
const status = document.querySelector('#status')
const subjectElement = document.querySelector('#subject')

function bodyText(item) {
  return new Promise(resolve => item.body.getAsync(Office.CoercionType.Text, result => resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : '')))
}

function updateButton() {
  button.disabled = keyInput.value.trim().length < 16
}

Office.onReady(() => {
  const item = Office.context.mailbox.item
  subjectElement.textContent = item.subject || 'Без темы'
  keyInput.value = localStorage.getItem('myplanner:pairing-key') || ''
  updateButton()
  keyInput.addEventListener('input', updateButton)
  button.addEventListener('click', async () => {
    const key = keyInput.value.trim()
    localStorage.setItem('myplanner:pairing-key', key)
    status.textContent = 'Передаю письмо…'
    const target = window.open('about:blank', '_blank')
    try {
      const text = (await bodyText(item)).replace(/\s+/g, ' ').trim().slice(0, 1200)
      const params = new URLSearchParams({
        key,
        itemId: item.itemId || '',
        subject: item.subject || '',
        senderName: item.from?.displayName || '',
        senderEmail: item.from?.emailAddress || '',
        receivedAt: item.dateTimeCreated instanceof Date ? item.dateTimeCreated.toISOString() : '',
        excerpt: text,
      })
      if (!target) throw new Error('Outlook заблокировал открытие локального окна.')
      target.location.href = `${bridgeUrl}?${params}`
      status.textContent = 'Письмо передано. Проверьте «Входящие» в MyPlanner.'
    } catch (error) {
      target?.close()
      status.textContent = error instanceof Error ? error.message : String(error)
    }
  })
})
