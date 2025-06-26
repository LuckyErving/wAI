const messagesEl = document.getElementById('messages');
const inputEl = document.getElementById('input');
const sendBtn = document.getElementById('send');
const loadingEl = document.getElementById('loading');

let chatHistory = []; // {role: 'user'|'ai', content: string}

function renderMessages() {
  messagesEl.innerHTML = '';
  chatHistory.forEach(msg => {
    const msgDiv = document.createElement('div');
    msgDiv.className = 'msg ' + (msg.role === 'user' ? 'user' : 'ai');
    const bubble = document.createElement('div');
    bubble.className = 'bubble';
    bubble.textContent = msg.content;
    msgDiv.appendChild(bubble);
    messagesEl.appendChild(msgDiv);
  });
  messagesEl.scrollTop = messagesEl.scrollHeight;
}

async function sendMessage() {
  const text = inputEl.value.trim();
  if (!text) return;
  chatHistory.push({ role: 'user', content: text });
  renderMessages();
  inputEl.value = '';
  setInputDisabled(true);
  loadingEl.style.display = '';
  try {
    const aiReply = await callAIChatAPI([...chatHistory]);
    chatHistory.push({ role: 'ai', content: aiReply });
    renderMessages();
  } catch (e) {
    chatHistory.push({ role: 'ai', content: '[AI接口出错] ' + e.message });
    renderMessages();
  } finally {
    setInputDisabled(false);
    loadingEl.style.display = 'none';
    inputEl.focus();
  }
}

function setInputDisabled(disabled) {
  inputEl.disabled = disabled;
  sendBtn.disabled = disabled;
}

sendBtn.onclick = sendMessage;
inputEl.onkeydown = e => {
  if (e.key === 'Enter' && !e.shiftKey) {
    e.preventDefault();
    sendMessage();
  }
};

function callAIChatAPI(history) {
  // 组装 messages
  const apiKey = "sk-wgR14x8Ec0njyIfT22Ab2554662a494d8d18807b57200686"; // 替换为你的API KEY
  const apiUrl = "https://free.v36.cm/v1/chat/completions";
  const messages = history.map(msg => ({
    role: msg.role === 'user' ? 'user' : 'assistant',
    content: msg.content
  }));
  return fetch(apiUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model: "gpt-3.5-turbo-1106",
      messages,
      max_tokens: 512
    })
  })
    .then(async res => {
      if (!res.ok) {
        const errText = await res.text();
        throw new Error(`AI接口请求失败，状态码: ${res.status}，返回内容: ${errText}`);
      }
      return res.json();
    })
    .then(data => data.choices[0].message.content.trim());
}
