// 视图 4：系统设置与模型 API 配置 (Phosphor Icons & 工业控制中心)

const PRESETS = {
  'deepseek': {
    baseUrl: 'https://api.deepseek.com/v1',
    modelName: 'deepseek-chat',
    desc: 'DeepSeek 官方开放平台 (推荐 · 极高性价比与制造业理解力)'
  },
  'gemini': {
    baseUrl: 'https://generativelanguage.googleapis.com/v1beta/openai',
    modelName: 'gemini-3.8-flash',
    desc: 'Google Gemini 官方平台 (默认推荐 gemini-3.8-flash · 极高推理速度与精度)'
  },
  'siliconflow': {
    baseUrl: 'https://api.siliconflow.cn/v1',
    modelName: 'deepseek-ai/DeepSeek-V3',
    desc: '硅基流动 (SiliconFlow) - 支持 DeepSeek-V3 / R1'
  },
  'dashscope': {
    baseUrl: 'https://dashscope.aliyuncs.com/compatible-mode/v1',
    modelName: 'deepseek-v3',
    desc: '阿里云百炼 (DashScope) - 支持 DeepSeek-V3 / 通义千问'
  },
  'zhipu': {
    baseUrl: 'https://open.bigmodel.cn/api/paas/v4',
    modelName: 'glm-4-flash',
    desc: '智谱开放平台 (GLM-4-Flash 免费高可用)'
  },
  'moonshot': {
    baseUrl: 'https://api.moonshot.cn/v1',
    modelName: 'moonshot-v1-8k',
    desc: '月之暗面 (Moonshot / Kimi)'
  },
  'openai': {
    baseUrl: 'https://api.openai.com/v1',
    modelName: 'gpt-4o-mini',
    desc: 'OpenAI 官方 (需境外网络或海外代理支持)'
  },
  'ollama': {
    baseUrl: 'http://localhost:11434/v1',
    modelName: 'qwen2.5:7b',
    desc: '本地离线 Ollama 模型服务 (无需联网与 Key)'
  },
  'custom': {
    baseUrl: '',
    modelName: '',
    desc: '自定义兼容接口 / OneAPI / 局域网网关'
  }
};

async function loadSettings() {
  try {
    const res = await fetch('/api/settings');
    const json = await res.json();
    if (json.status === 'success') {
      const data = json.data;
      AppState.settings = data;

      const providerSelect = document.getElementById('settings-provider-select');
      if (providerSelect) {
        providerSelect.value = data.provider || 'deepseek';
      }

      document.getElementById('settings-api-key').value = data.api_key || '';
      document.getElementById('settings-base-url').value = data.base_url || 'https://api.deepseek.com/v1';
      document.getElementById('settings-model-name').value = data.model_name || 'deepseek-chat';
      
      const exportDirInput = document.getElementById('settings-export-dir');
      if (exportDirInput) {
        exportDirInput.value = data.export_dir || '';
        // 如果未设置自定义导出路径，获取默认路径显示为 placeholder
        if (!data.export_dir) {
          fetch('/api/settings/default-export-dir')
            .then(r => r.json())
            .then(d => {
              if (d.status === 'success' && d.default_export_dir) {
                exportDirInput.placeholder = d.default_export_dir;
              }
            })
            .catch(() => {});
        }
      }

      // 同步设置面板中的主题卡片高亮状态
      updateSettingsThemeCards(AppState.currentTheme || data.theme || 'dark');
    }
  } catch (err) {
    console.error('获取设置失败:', err);
  }
}

// 浏览并选择本地导出目录
async function browseExportDir() {
  const currentVal = document.getElementById('settings-export-dir').value.trim();
  try {
    const res = await fetch('/api/settings/browse-export-dir', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ initial_dir: currentVal || '' })
    });
    const json = await res.json();
    if (json.status === 'success' && json.selected_dir) {
      document.getElementById('settings-export-dir').value = json.selected_dir;
      showToast(`已选择导出目录: ${json.selected_dir}`, 'success');
    } else if (json.status === 'cancelled') {
      // 用户取消选取，不报错
    }
  } catch (err) {
    showToast('选择目录异常: ' + err.message, 'error');
  }
}

// 在资源管理器中打开导出目录
async function openExportDir() {
  const currentVal = document.getElementById('settings-export-dir').value.trim();
  try {
    const res = await fetch('/api/settings/open-export-dir', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ export_dir: currentVal || '' })
    });
    const json = await res.json();
    if (json.status === 'success') {
      showToast(json.message || '已在资源管理器中打开导出目录', 'success');
    } else {
      showToast(json.detail || '打开导出目录失败', 'error');
    }
  } catch (err) {
    showToast('打开目录异常: ' + err.message, 'error');
  }
}

// 恢复系统默认导出目录
async function resetExportDir() {
  try {
    const res = await fetch('/api/settings/default-export-dir');
    const json = await res.json();
    if (json.status === 'success' && json.default_export_dir) {
      document.getElementById('settings-export-dir').value = json.default_export_dir;
      showToast('已重置为系统默认导出路径: ' + json.default_export_dir, 'info');
    }
  } catch (err) {
    showToast('恢复默认路径失败: ' + err.message, 'error');
  }
}

// 在设置面板中切换主题卡片
function selectThemeInSettings(themeKey) {
  applyTheme(themeKey, true);
  saveOptionToBackend({ theme: themeKey });
  updateSettingsThemeCards(themeKey);
}

function handleProviderPresetChange(providerKey) {
  const preset = PRESETS[providerKey];
  if (!preset) return;

  const baseUrlInput = document.getElementById('settings-base-url');
  const modelNameInput = document.getElementById('settings-model-name');
  const presetTip = document.getElementById('settings-preset-tip');

  if (preset.baseUrl) baseUrlInput.value = preset.baseUrl;
  if (preset.modelName) modelNameInput.value = preset.modelName;
  if (presetTip) presetTip.innerText = preset.desc;

  const keyInput = document.getElementById('settings-api-key');
  if (keyInput) {
    if (providerKey === 'gemini') {
      keyInput.placeholder = '填入 Google AI Studio 的 Gemini API Key (AIzaSy...)';
    } else if (providerKey === 'ollama') {
      keyInput.placeholder = 'Ollama 本地服务无需配置 API Key';
    } else {
      keyInput.placeholder = '填入 sk- 开头的 API Key 密钥';
    }
  }
}

function toggleApiKeyVisibility() {
  const input = document.getElementById('settings-api-key');
  const icon = document.getElementById('toggle-key-icon');
  if (input.type === 'password') {
    input.type = 'text';
    icon.className = 'ph ph-eye-slash text-base';
  } else {
    input.type = 'password';
    icon.className = 'ph ph-eye text-base';
  }
}

async function testApiConnection() {
  const provider = document.getElementById('settings-provider-select')?.value || 'deepseek';
  const apiKey = document.getElementById('settings-api-key').value.trim();
  const baseUrl = document.getElementById('settings-base-url').value.trim();
  const modelName = document.getElementById('settings-model-name').value.trim();

  if (!apiKey && provider !== 'ollama') {
    showToast('请输入 API Key 后再进行连通性测试', 'error');
    return;
  }

  const btnTest = document.getElementById('btn-test-conn');
  if (btnTest) {
    btnTest.disabled = true;
    btnTest.innerHTML = '<i class="ph ph-spinner animate-spin text-sm"></i> 正在测试...';
  }

  showToast('正在向大模型 API 发起连通性握手测试...', 'info');

  try {
    const res = await fetch('/api/settings/test-connection', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        api_key: apiKey,
        base_url: baseUrl,
        model_name: modelName
      })
    });
    const json = await res.json();

    if (json.success) {
      showToast(json.message, 'success');
      
      // 测试成功后自动静默保存配置并更新侧边栏状态
      await saveAllSettings(true);
      updateEngineStatusUI({
        success: true,
        model_name: modelName,
        message: json.message
      });
    } else {
      showToast(json.message, 'error');
      updateEngineStatusUI({
        success: false,
        model_name: modelName,
        message: json.message
      });
      alert(`【API 连接测试失败】\n\n${json.message}`);
    }
  } catch (err) {
    showToast('连接测试异常: ' + err.message, 'error');
  } finally {
    if (btnTest) {
      btnTest.disabled = false;
      btnTest.innerHTML = '<i class="ph ph-plugs-connected text-sm"></i> 测试连通性';
    }
  }
}

async function saveAllSettings(silent = false) {
  const provider = document.getElementById('settings-provider-select')?.value || 'deepseek';
  const apiKey = document.getElementById('settings-api-key').value.trim();
  const baseUrl = document.getElementById('settings-base-url').value.trim();
  const modelName = document.getElementById('settings-model-name').value.trim();
  const exportDir = document.getElementById('settings-export-dir').value.trim();
  const theme = AppState.currentTheme || 'dark';

  try {
    const res = await fetch('/api/settings', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        provider: provider,
        api_key: apiKey,
        base_url: baseUrl,
        model_name: modelName,
        export_dir: exportDir,
        theme: theme
      })
    });
    const json = await res.json();

    if (json.status === 'success') {
      if (!silent) showToast('系统与模型设置已成功保存！', 'success');
      loadInitialSettings();
    } else {
      if (!silent) showToast(json.detail || '保存失败', 'error');
    }
  } catch (err) {
    if (!silent) showToast('保存异常: ' + err.message, 'error');
  }
}
