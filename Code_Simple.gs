/**
 * Google Apps Script for Google Sheets Translation Tool - Phiên bản đơn giản
 * Không sử dụng ui.alert() để tránh lỗi
 */

// Configuration
const GEMINI_API_KEY = 'YOUR_GEMINI_API_KEY_HERE';
const OPENAI_API_KEY = 'YOUR_OPENAI_API_KEY_HERE';
const AVAILABLE_MODELS = {
  'gemini-2.0-flash-lite': {
    name: 'Gemini 2.0 Flash Lite',
    url: 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-lite:generateContent',
    cost: 'Cao',
    quality: 'Tốt nhất',
    provider: 'Google'
  },
  'gemma-3-4b-it': {
    name: 'Gemma 3 4B IT',
    url: 'https://generativelanguage.googleapis.com/v1beta/models/gemma-3-4b-it:generateContent',
    cost: 'Thấp',
    quality: 'Tốt',
    provider: 'Google'
  },
  'gpt-4o': {
    name: 'GPT-4o',
    url: 'https://api.openai.com/v1/chat/completions',
    cost: 'Cao',
    quality: 'Tốt nhất',
    provider: 'OpenAI'
  },
  'gpt-4o-mini': {
    name: 'GPT-4o Mini',
    url: 'https://api.openai.com/v1/chat/completions',
    cost: 'Trung bình',
    quality: 'Tốt',
    provider: 'OpenAI'
  },
  'gpt-3.5-turbo': {
    name: 'GPT-3.5 Turbo',
    url: 'https://api.openai.com/v1/chat/completions',
    cost: 'Thấp',
    quality: 'Khá',
    provider: 'OpenAI'
  }
};

/**
 * Tạo menu trong Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔄 Translation Tool')
    .addItem('Dịch vùng chọn', 'translateSelectedRange')
    .addSeparator()
    .addItem('Cài đặt API Key', 'showApiKeyDialog')
    .addItem('Cài đặt Model', 'showModelDialog')
    .addItem('🔒 Bảo mật API Keys', 'cleanupApiKeys')
    .addItem('🔍 Kiểm tra bảo mật', 'checkApiKeySecurity')
    .addToUi();
}

/**
 * Hiển thị dialog để nhập API Key
 */
function showApiKeyDialog() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      .input-group { margin-bottom: 15px; }
      label { display: block; margin-bottom: 5px; font-weight: bold; }
      input[type="text"] { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; }
      button { background: #4285f4; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; }
      button:hover { background: #3367d6; }
      .api-section { 
        border: 1px solid #e1e5e9; 
        border-radius: 8px; 
        padding: 15px; 
        margin-bottom: 15px; 
      }
      .api-title { font-weight: bold; font-size: 16px; margin-bottom: 10px; }
      .api-description { color: #666; font-size: 14px; margin-bottom: 10px; }
    </style>
    <h3>🔑 Cài đặt API Keys</h3>
    
    <div class="api-section">
      <div class="api-title">🤖 Google Gemini API</div>
      <div class="api-description">Sử dụng cho Gemini và Gemma models</div>
      <div class="input-group">
        <label for="geminiApiKey">Gemini API Key:</label>
        <input type="text" id="geminiApiKey" placeholder="Nhập Gemini API Key của bạn">
      </div>
    </div>
    
    <div class="api-section">
      <div class="api-title">🧠 OpenAI GPT API</div>
      <div class="api-description">Sử dụng cho GPT models</div>
      <div class="input-group">
        <label for="openaiApiKey">OpenAI API Key:</label>
        <input type="text" id="openaiApiKey" placeholder="Nhập OpenAI API Key của bạn">
      </div>
    </div>
    
    <button onclick="saveApiKeys()">Lưu API Keys</button>
    <script>
      function saveApiKeys() {
        const geminiApiKey = document.getElementById('geminiApiKey').value;
        const openaiApiKey = document.getElementById('openaiApiKey').value;
        
        if (geminiApiKey.trim() === '' && openaiApiKey.trim() === '') {
          alert('Vui lòng nhập ít nhất một API Key!');
          return;
        }
        
        google.script.run
          .withSuccessHandler(function() {
            alert('API Keys đã được lưu thành công!');
            google.script.host.close();
          })
          .withFailureHandler(function(error) {
            alert('Lỗi: ' + error);
          })
          .saveApiKeys(geminiApiKey, openaiApiKey);
      }
    </script>
  `)
    .setWidth(500)
    .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Cài đặt API Keys');
}

/**
 * Hiển thị dialog để chọn model
 */
function showModelDialog() {
  const currentModel = getSelectedModel();
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      .model-option { 
        border: 2px solid #e1e5e9; 
        border-radius: 8px; 
        padding: 15px; 
        margin-bottom: 10px; 
        cursor: pointer;
        transition: all 0.3s ease;
      }
      .model-option:hover { border-color: #4285f4; background: #f8f9fa; }
      .model-option.selected { border-color: #4285f4; background: #e8f4fd; }
      .model-name { font-weight: bold; font-size: 16px; margin-bottom: 5px; }
      .model-details { color: #666; font-size: 14px; }
      .cost-badge { 
        display: inline-block; 
        padding: 2px 8px; 
        border-radius: 12px; 
        font-size: 12px; 
        font-weight: bold;
        margin-left: 10px;
      }
      .cost-high { background: #ffebee; color: #c62828; }
      .cost-medium { background: #fff3e0; color: #ef6c00; }
      .cost-low { background: #e8f5e8; color: #2e7d32; }
      .provider-badge {
        display: inline-block;
        padding: 2px 6px;
        border-radius: 10px;
        font-size: 10px;
        font-weight: bold;
        margin-left: 5px;
      }
      .provider-google { background: #e3f2fd; color: #1976d2; }
      .provider-openai { background: #f3e5f5; color: #7b1fa2; }
      button { background: #4285f4; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin-top: 15px; }
      button:hover { background: #3367d6; }
    </style>
    <h3>🤖 Chọn Model AI</h3>
    <p>Chọn model phù hợp với nhu cầu và ngân sách:</p>
    
    <div class="model-option ${currentModel === 'gpt-4o' ? 'selected' : ''}" onclick="selectModel('gpt-4o')">
      <div class="model-name">
        GPT-4o
        <span class="cost-badge cost-high">Chi phí cao</span>
        <span class="provider-badge provider-openai">OpenAI</span>
      </div>
      <div class="model-details">Chất lượng dịch tốt nhất, phù hợp cho nội dung quan trọng</div>
    </div>
    
    <div class="model-option ${currentModel === 'gpt-4o-mini' ? 'selected' : ''}" onclick="selectModel('gpt-4o-mini')">
      <div class="model-name">
        GPT-4o Mini
        <span class="cost-badge cost-medium">Chi phí trung bình</span>
        <span class="provider-badge provider-openai">OpenAI</span>
      </div>
      <div class="model-details">Cân bằng giữa chất lượng và chi phí</div>
    </div>
    
    <div class="model-option ${currentModel === 'gpt-3.5-turbo' ? 'selected' : ''}" onclick="selectModel('gpt-3.5-turbo')">
      <div class="model-name">
        GPT-3.5 Turbo
        <span class="cost-badge cost-low">Tiết kiệm</span>
        <span class="provider-badge provider-openai">OpenAI</span>
      </div>
      <div class="model-details">Chi phí thấp, phù hợp cho dịch thuật thường xuyên</div>
    </div>
    
    <div class="model-option ${currentModel === 'gemini-2.0-flash-lite' ? 'selected' : ''}" onclick="selectModel('gemini-2.0-flash-lite')">
      <div class="model-name">
        Gemini 2.0 Flash Lite
        <span class="cost-badge cost-high">Chi phí cao</span>
        <span class="provider-badge provider-google">Google</span>
      </div>
      <div class="model-details">Chất lượng dịch tốt nhất từ Google</div>
    </div>
    
    <div class="model-option ${currentModel === 'gemma-3-4b-it' ? 'selected' : ''}" onclick="selectModel('gemma-3-4b-it')">
      <div class="model-name">
        Gemma 3 4B IT
        <span class="cost-badge cost-low">Tiết kiệm</span>
        <span class="provider-badge provider-google">Google</span>
      </div>
      <div class="model-details">Chi phí thấp, chất lượng tốt từ Google</div>
    </div>
    
    <button onclick="saveModel()">Lưu lựa chọn</button>
    
    <script>
      let selectedModel = '${currentModel}';
      
      function selectModel(model) {
        document.querySelectorAll('.model-option').forEach(option => {
          option.classList.remove('selected');
        });
        event.currentTarget.classList.add('selected');
        selectedModel = model;
      }
      
      function saveModel() {
        google.script.run
          .withSuccessHandler(function() {
            alert('Model đã được lưu thành công!');
            google.script.host.close();
          })
          .withFailureHandler(function(error) {
            alert('Lỗi: ' + error);
          })
          .saveSelectedModel(selectedModel);
      }
    </script>
  `)
    .setWidth(600)
    .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Chọn Model AI');
}

/**
 * Lưu API Keys vào Properties
 */
function saveApiKeys(geminiApiKey, openaiApiKey) {
  if (geminiApiKey && geminiApiKey.trim() !== '') {
    PropertiesService.getUserProperties().setProperty('GEMINI_API_KEY', geminiApiKey);
  }
  if (openaiApiKey && openaiApiKey.trim() !== '') {
    PropertiesService.getUserProperties().setProperty('OPENAI_API_KEY', openaiApiKey);
  }
}

/**
 * Lưu API Key vào Properties (backward compatibility)
 */
function saveApiKey(apiKey) {
  PropertiesService.getUserProperties().setProperty('GEMINI_API_KEY', apiKey);
}

/**
 * Lưu model được chọn vào Properties
 */
function saveSelectedModel(model) {
  PropertiesService.getUserProperties().setProperty('SELECTED_MODEL', model);
}

/**
 * Lấy API Key từ Properties
 */
function getApiKey() {
  const selectedModel = getSelectedModel();
  const modelInfo = AVAILABLE_MODELS[selectedModel];
  
  if (modelInfo.provider === 'OpenAI') {
    const apiKey = PropertiesService.getUserProperties().getProperty('OPENAI_API_KEY');
    if (!apiKey) {
      throw new Error('OpenAI API Key chưa được cài đặt. Vui lòng sử dụng menu "Cài đặt API Key"');
    }
    return apiKey;
  } else {
    // Google models
    const apiKey = PropertiesService.getUserProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) {
      throw new Error('Gemini API Key chưa được cài đặt. Vui lòng sử dụng menu "Cài đặt API Key"');
    }
    return apiKey;
  }
}

/**
 * Lấy model được chọn từ Properties
 */
function getSelectedModel() {
  const model = PropertiesService.getUserProperties().getProperty('SELECTED_MODEL');
  return model || 'gemma-3-4b-it';
}

/**
 * Dọn dẹp và bảo mật API Keys - Di chuyển từ Script Properties sang User Properties
 */
function cleanupApiKeys() {
  console.log('🔒 Bắt đầu dọn dẹp API Keys để bảo mật...');
  
  // Kiểm tra xem có API keys cũ trong Script Properties không
  const scriptGeminiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const scriptOpenaiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  const scriptModel = PropertiesService.getScriptProperties().getProperty('SELECTED_MODEL');
  
  // Kiểm tra User Properties hiện tại
  const userGeminiKey = PropertiesService.getUserProperties().getProperty('GEMINI_API_KEY');
  const userOpenaiKey = PropertiesService.getUserProperties().getProperty('OPENAI_API_KEY');
  const userModel = PropertiesService.getUserProperties().getProperty('SELECTED_MODEL');
  
  let migrated = false;
  
  // Di chuyển API keys từ Script sang User Properties nếu cần
  if (scriptGeminiKey && !userGeminiKey) {
    PropertiesService.getUserProperties().setProperty('GEMINI_API_KEY', scriptGeminiKey);
    console.log('✅ Đã di chuyển Gemini API Key sang User Properties');
    migrated = true;
  }
  
  if (scriptOpenaiKey && !userOpenaiKey) {
    PropertiesService.getUserProperties().setProperty('OPENAI_API_KEY', scriptOpenaiKey);
    console.log('✅ Đã di chuyển OpenAI API Key sang User Properties');
    migrated = true;
  }
  
  if (scriptModel && !userModel) {
    PropertiesService.getUserProperties().setProperty('SELECTED_MODEL', scriptModel);
    console.log('✅ Đã di chuyển Model selection sang User Properties');
    migrated = true;
  }
  
  // Xóa API keys cũ khỏi Script Properties để bảo mật
  if (scriptGeminiKey || scriptOpenaiKey || scriptModel) {
    PropertiesService.getScriptProperties().deleteProperty('GEMINI_API_KEY');
    PropertiesService.getScriptProperties().deleteProperty('OPENAI_API_KEY');
    PropertiesService.getScriptProperties().deleteProperty('SELECTED_MODEL');
    console.log('🗑️ Đã xóa API Keys cũ khỏi Script Properties');
  }
  
  if (migrated) {
    console.log('🎉 Hoàn thành di chuyển API Keys sang User Properties');
    console.log('🔒 API Keys của bạn giờ đây được bảo vệ và chỉ bạn mới có thể truy cập!');
  } else {
    console.log('ℹ️ Không có API Keys cũ cần di chuyển');
  }
  
  // Hiển thị trạng thái hiện tại
  console.log('\n📊 Trạng thái API Keys hiện tại:');
  console.log(`Gemini API Key: ${userGeminiKey ? '✅ Đã cài đặt' : '❌ Chưa cài đặt'}`);
  console.log(`OpenAI API Key: ${userOpenaiKey ? '✅ Đã cài đặt' : '❌ Chưa cài đặt'}`);
  console.log(`Selected Model: ${userModel || 'gemma-3-4b-it'}`);
}

/**
 * Kiểm tra bảo mật API Keys
 */
function checkApiKeySecurity() {
  console.log('🔒 === KIỂM TRA BẢO MẬT API KEYS ===');
  
  // Kiểm tra Script Properties (không an toàn)
  const scriptGeminiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const scriptOpenaiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  
  // Kiểm tra User Properties (an toàn)
  const userGeminiKey = PropertiesService.getUserProperties().getProperty('GEMINI_API_KEY');
  const userOpenaiKey = PropertiesService.getUserProperties().getProperty('OPENAI_API_KEY');
  
  let securityIssues = [];
  let securityGood = [];
  
  // Kiểm tra Gemini API Key
  if (scriptGeminiKey) {
    securityIssues.push('❌ Gemini API Key đang lưu trong Script Properties (KHÔNG AN TOÀN)');
  } else if (userGeminiKey) {
    securityGood.push('✅ Gemini API Key được lưu trong User Properties (AN TOÀN)');
  } else {
    securityGood.push('ℹ️ Chưa cài đặt Gemini API Key');
  }
  
  // Kiểm tra OpenAI API Key
  if (scriptOpenaiKey) {
    securityIssues.push('❌ OpenAI API Key đang lưu trong Script Properties (KHÔNG AN TOÀN)');
  } else if (userOpenaiKey) {
    securityGood.push('✅ OpenAI API Key được lưu trong User Properties (AN TOÀN)');
  } else {
    securityGood.push('ℹ️ Chưa cài đặt OpenAI API Key');
  }
  
  // Hiển thị kết quả
  if (securityIssues.length > 0) {
    console.log('⚠️ PHÁT HIỆN VẤN ĐỀ BẢO MẬT:');
    securityIssues.forEach(issue => console.log(issue));
    console.log('\n💡 GIẢI PHÁP: Sử dụng menu "🔒 Bảo mật API Keys" để di chuyển API Keys sang User Properties');
  } else {
    console.log('🎉 TẤT CẢ API KEYS ĐÃ ĐƯỢC BẢO VỆ!');
  }
  
  securityGood.forEach(good => console.log(good));
  
  // Tóm tắt
  console.log('\n📊 TÓM TẮT BẢO MẬT:');
  console.log(`Script Properties (không an toàn): ${scriptGeminiKey || scriptOpenaiKey ? '❌ Có API Keys' : '✅ Không có API Keys'}`);
  console.log(`User Properties (an toàn): ${userGeminiKey || userOpenaiKey ? '✅ Có API Keys' : 'ℹ️ Chưa có API Keys'}`);
  
  if (securityIssues.length === 0) {
    console.log('🔒 API Keys của bạn đã được bảo vệ hoàn toàn!');
    console.log('👥 Người khác không thể truy cập API Keys của bạn.');
  } else {
    console.log('🚨 Cần thực hiện bảo mật ngay!');
  }
}

/**
 * Lấy API URL dựa trên model được chọn
 */
function getApiUrl() {
  const model = getSelectedModel();
  return AVAILABLE_MODELS[model]?.url || AVAILABLE_MODELS['gemma-3-4b-it'].url;
}

/**
 * Dịch text sử dụng Gemini/Gemma/GPT API với retry logic
 */
function translateText(text, targetLanguage = 'Vietnamese') {
  if (!text || text.toString().trim() === '') {
    return '';
  }
  
  const apiKey = getApiKey();
  const apiUrl = getApiUrl();
  const selectedModel = getSelectedModel();
  const modelInfo = AVAILABLE_MODELS[selectedModel];
  
  let prompt;
  let payload;
  let options;
  
  if (modelInfo.provider === 'OpenAI') {
    // OpenAI GPT models
    prompt = `Translate the following text to ${targetLanguage}. Return only the translation without any additional text or formatting: "${text}"`;
    
    payload = {
      model: selectedModel,
      messages: [
        {
          role: "user",
          content: prompt
        }
      ],
      temperature: 0.3,
      max_tokens: 1000
    };
    
    options = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
  } else {
    // Google Gemini/Gemma models
    if (selectedModel === 'gemma-3-4b-it') {
      prompt = `Translate this text to ${targetLanguage}. Return only the translation: "${text}"`;
    } else {
      prompt = `Dịch đoạn text sau sang tiếng ${targetLanguage}. Chỉ trả về bản dịch, không thêm giải thích hay dấu ngoặc kép: ${text}`;
    }

    payload = {
      contents: [{
        parts: [{
          text: prompt
        }]
      }]
    };

    options = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-goog-api-key': apiKey
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
  }

  // Retry logic - thử tối đa 3 lần
  const maxRetries = 3;
  let lastError = null;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const response = UrlFetchApp.fetch(apiUrl, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      // Kiểm tra HTTP status code
      if (responseCode !== 200) {
        throw new Error(`HTTP ${responseCode}: ${responseText}`);
      }
      
      const data = JSON.parse(responseText);
      let translatedText = '';
      
      if (modelInfo.provider === 'OpenAI') {
        // Parse OpenAI response
        if (data.choices && data.choices[0] && data.choices[0].message) {
          translatedText = data.choices[0].message.content.trim();
        } else {
          throw new Error('Không nhận được phản hồi hợp lệ từ OpenAI API');
        }
      } else {
        // Parse Google response
        if (data.candidates && data.candidates[0] && data.candidates[0].content) {
          translatedText = data.candidates[0].content.parts[0].text.trim();
        } else {
          throw new Error('Không nhận được phản hồi hợp lệ từ Google API');
        }
      }
      
      // Kiểm tra kết quả dịch có hợp lệ không
      if (translatedText && translatedText.length > 0) {
        return translatedText;
      } else {
        throw new Error('Kết quả dịch rỗng');
      }
      
    } catch (error) {
      lastError = error;
      console.log(`Lần thử ${attempt}/${maxRetries} thất bại: ${error.message}`);
      
      // Nếu còn lần thử, chờ một lúc rồi thử lại
      if (attempt < maxRetries) {
        const delay = selectedModel === 'gemma-3-4b-it' ? 1000 : 1500;
        console.log(`Chờ ${delay}ms trước khi thử lại...`);
        Utilities.sleep(delay);
      }
    }
  }
  
  // Nếu tất cả lần thử đều thất bại
  console.error(`Dịch thất bại sau ${maxRetries} lần thử: ${lastError.message}`);
  throw new Error(`Dịch thất bại: ${lastError.message}`);
}

/**
 * Dịch chunk text (nhiều text cùng lúc) với separator "|||"
 */
function translateTextChunk(texts, targetLanguage = 'Vietnamese') {
  if (!texts || texts.length === 0) {
    return [];
  }
  
  const apiKey = getApiKey();
  const apiUrl = getApiUrl();
  const selectedModel = getSelectedModel();
  const modelInfo = AVAILABLE_MODELS[selectedModel];
  
  // Tạo prompt cho chunk
  let prompt;
  if (modelInfo.provider === 'OpenAI') {
    // OpenAI GPT models - sử dụng "|||" separator
    prompt = `Translate the following texts to ${targetLanguage}. Each text may contain multiple lines or paragraphs - treat each numbered item as ONE complete text. Return only the translations, separated by "|||":\n\n`;
    for (let i = 0; i < texts.length; i++) {
      prompt += `${i + 1}. ${texts[i]}\n\n`;
    }
  } else {
    // Google models - sử dụng "|||" separator
    if (selectedModel === 'gemma-3-4b-it') {
      prompt = `Translate the following texts to ${targetLanguage}. Each text may contain multiple lines - treat each numbered item as ONE complete text. Return only the translations, separated by "|||":\n\n`;
    } else {
      prompt = `Dịch các đoạn text sau sang tiếng ${targetLanguage}. Mỗi đoạn text có thể có nhiều dòng - coi mỗi mục đánh số là MỘT văn bản hoàn chỉnh. Chỉ trả về bản dịch, phân cách bằng "|||":\n\n`;
    }
    
    for (let i = 0; i < texts.length; i++) {
      prompt += `${i + 1}. ${texts[i]}\n\n`;
    }
  }
  
  let payload;
  let options;
  
  if (modelInfo.provider === 'OpenAI') {
    payload = {
      model: selectedModel,
      messages: [
        {
          role: "user",
          content: prompt
        }
      ],
      temperature: 0.3,
      max_tokens: 2000
    };
    
    options = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
  } else {
    payload = {
      contents: [{
        parts: [{
          text: prompt
        }]
      }]
    };

    options = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-goog-api-key': apiKey
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
  }

  // Retry logic - thử tối đa 3 lần
  const maxRetries = 3;
  let lastError = null;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const response = UrlFetchApp.fetch(apiUrl, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      // Kiểm tra HTTP status code
      if (responseCode !== 200) {
        throw new Error(`HTTP ${responseCode}: ${responseText}`);
      }
      
      const data = JSON.parse(responseText);
      let translatedText = '';
      
      if (modelInfo.provider === 'OpenAI') {
        // Parse OpenAI response
        if (data.choices && data.choices[0] && data.choices[0].message) {
          translatedText = data.choices[0].message.content.trim();
        } else {
          throw new Error('Không nhận được phản hồi hợp lệ từ OpenAI API');
        }
      } else {
        // Parse Google response
        if (data.candidates && data.candidates[0] && data.candidates[0].content) {
          translatedText = data.candidates[0].content.parts[0].text.trim();
        } else {
          throw new Error('Không nhận được phản hồi hợp lệ từ Google API');
        }
      }
      
      // Tách kết quả bằng "|||" - tất cả models đều sử dụng cùng separator
      const translatedParts = translatedText.split('|||').map(part => part.trim()).filter(part => part !== '');
      
      // Kiểm tra số lượng kết quả
      if (translatedParts.length === 0) {
        throw new Error('Không nhận được kết quả dịch hợp lệ');
      }
      
      console.log(`✅ Chunk dịch thành công: ${texts.length} text -> ${translatedParts.length} kết quả`);
      
      // Xử lý trường hợp số kết quả không khớp với số input
      if (translatedParts.length !== texts.length) {
        console.log(`⚠️ Cảnh báo: Số kết quả (${translatedParts.length}) không khớp với số input (${texts.length})`);
        
        // Nếu có nhiều kết quả hơn input, ghép lại
        if (translatedParts.length > texts.length) {
          const adjustedResults = [];
          let currentIndex = 0;
          
          for (let i = 0; i < texts.length; i++) {
            if (currentIndex < translatedParts.length) {
              // Kiểm tra xem text gốc có xuống dòng không
              const originalText = texts[i];
              const lineCount = (originalText.match(/\n/g) || []).length + 1;
              
              if (lineCount > 1 && currentIndex + lineCount <= translatedParts.length) {
                // Ghép các kết quả tương ứng với số dòng
                const combinedResult = translatedParts.slice(currentIndex, currentIndex + lineCount).join('\n');
                adjustedResults.push(combinedResult);
                currentIndex += lineCount;
              } else {
                // Lấy kết quả đầu tiên
                adjustedResults.push(translatedParts[currentIndex]);
                currentIndex++;
              }
            } else {
              // Không đủ kết quả, giữ nguyên text gốc
              adjustedResults.push(texts[i]);
            }
          }
          
          console.log(`🔄 Đã điều chỉnh kết quả: ${translatedParts.length} -> ${adjustedResults.length}`);
          return adjustedResults;
        }
        
        // Nếu có ít kết quả hơn input, bổ sung text gốc
        if (translatedParts.length < texts.length) {
          const adjustedResults = [...translatedParts];
          for (let i = translatedParts.length; i < texts.length; i++) {
            adjustedResults.push(texts[i]); // Giữ nguyên text gốc
          }
          console.log(`🔄 Đã bổ sung text gốc cho các kết quả thiếu`);
          return adjustedResults;
        }
      }
      
      return translatedParts;
      
    } catch (error) {
      lastError = error;
      console.log(`Lần thử ${attempt}/${maxRetries} thất bại: ${error.message}`);
      
      // Nếu còn lần thử, chờ một lúc rồi thử lại
      if (attempt < maxRetries) {
        const delay = selectedModel === 'gemma-3-4b-it' ? 2000 : 2500;
        console.log(`Chờ ${delay}ms trước khi thử lại...`);
        Utilities.sleep(delay);
      }
    }
  }
  
  // Nếu tất cả lần thử đều thất bại
  console.error(`Dịch chunk thất bại sau ${maxRetries} lần thử: ${lastError.message}`);
  throw new Error(`Dịch chunk thất bại: ${lastError.message}`);
}

/**
 * Tạo tên sheet dịch
 */
function getTranslationSheetName(originalSheetName, targetLanguage) {
  const languageSuffix = targetLanguage.toLowerCase();
  return `${originalSheetName}_${languageSuffix}`;
}

/**
 * Tạo hoặc lấy sheet dịch
 */
function getOrCreateTranslationSheet(spreadsheet, originalSheetName, targetLanguage) {
  const translationSheetName = getTranslationSheetName(originalSheetName, targetLanguage);
  
  let translationSheet = spreadsheet.getSheetByName(translationSheetName);
  
  if (translationSheet) {
    translationSheet.clear();
  } else {
    translationSheet = spreadsheet.insertSheet(translationSheetName);
  }
  
  return translationSheet;
}

/**
 * Copy format và style từ sheet gốc
 */
function copySheetFormat(sourceSheet, targetSheet) {
  try {
    const numCols = sourceSheet.getLastColumn();
    const numRows = sourceSheet.getLastRow();
    
    for (let col = 1; col <= numCols; col++) {
      const width = sourceSheet.getColumnWidth(col);
      targetSheet.setColumnWidth(col, width);
    }
    
    for (let row = 1; row <= numRows; row++) {
      const height = sourceSheet.getRowHeight(row);
      targetSheet.setRowHeight(row, height);
    }
    
    if (numRows > 0) {
      const headerRange = sourceSheet.getRange(1, 1, 1, numCols);
      const headerFormat = headerRange.getFontWeight();
      const headerBackground = headerRange.getBackground();
      
      const targetHeaderRange = targetSheet.getRange(1, 1, 1, numCols);
      targetHeaderRange.setFontWeight(headerFormat);
      targetHeaderRange.setBackground(headerBackground);
    }
  } catch (error) {
    console.log('Không thể copy format:', error.message);
  }
}

/**
 * Dịch vùng được chọn trong sheet
 */
function translateSelectedRange() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  
  if (!range) {
    console.log('❌ Vui lòng chọn một vùng để dịch!');
    return;
  }
  
  const selectedModel = getSelectedModel();
  const modelInfo = AVAILABLE_MODELS[selectedModel];
  
  console.log(`🔄 Bắt đầu dịch vùng được chọn với ${modelInfo.name}`);
  console.log(`📍 Vùng: ${range.getA1Notation()}`);
  
  const targetLanguage = 'Vietnamese';
  
  try {
    translateRangeInPlace(range, targetLanguage);
  } catch (error) {
    console.error('Lỗi khi dịch vùng được chọn:', error);
  }
}

/**
 * Dịch vùng cụ thể tại chỗ - Phiên bản tối ưu với chunk processing
 */
function translateRangeInPlace(range, targetLanguage) {
  const values = range.getValues();
  const numRows = values.length;
  const numCols = values[0].length;
  
  let translatedCount = 0;
  let errorCount = 0;
  let emptyCount = 0;
  let totalCells = numRows * numCols;
  
  const selectedModel = getSelectedModel();
  const modelInfo = AVAILABLE_MODELS[selectedModel];
  
  console.log(`📊 Phân tích vùng được chọn`);
  console.log(`📝 Tổng số ô: ${totalCells}`);
  console.log(`🤖 Model: ${modelInfo.name}`);
  
  // Thu thập tất cả text cần dịch và vị trí
  const textsToTranslate = [];
  const cellPositions = [];
  
  for (let row = 0; row < numRows; row++) {
    for (let col = 0; col < numCols; col++) {
      const cellValue = values[row][col];
      if (cellValue && typeof cellValue === 'string' && cellValue.trim() !== '') {
        textsToTranslate.push(cellValue);
        cellPositions.push({ row, col });
      } else {
        emptyCount++;
      }
    }
  }
  
  console.log(`📝 Số ô có nội dung cần dịch: ${textsToTranslate.length}`);
  console.log(`⬜ Số ô rỗng: ${emptyCount}`);
  
  if (textsToTranslate.length === 0) {
    console.log(`ℹ️ Không có text nào cần dịch trong vùng được chọn!`);
    return;
  }
  
  // Chia thành các chunk (30 ô mỗi chunk)
  const CHUNK_SIZE = 30;
  const chunks = [];
  
  for (let i = 0; i < textsToTranslate.length; i += CHUNK_SIZE) {
    chunks.push(textsToTranslate.slice(i, i + CHUNK_SIZE));
  }
  
  console.log(`📦 Chia thành ${chunks.length} chunk (${CHUNK_SIZE} text/chunk)`);
  
  // Dịch từng chunk
  const allTranslatedTexts = [];
  
  for (let chunkIndex = 0; chunkIndex < chunks.length; chunkIndex++) {
    const chunk = chunks[chunkIndex];
    console.log(`\n🔄 Đang xử lý chunk ${chunkIndex + 1}/${chunks.length} (${chunk.length} text)`);
    
    try {
      // Sử dụng function translateTextChunk mới
      const translatedParts = translateTextChunk(chunk, targetLanguage);
      
      console.log(`✅ Chunk ${chunkIndex + 1} hoàn thành`);
      console.log(`📝 Số text gốc: ${chunk.length}`);
      console.log(`📝 Số text dịch: ${translatedParts.length}`);
      
      // Map kết quả về đúng vị trí
      for (let i = 0; i < chunk.length; i++) {
        const translatedText = translatedParts[i] || chunk[i]; // Giữ nguyên text gốc nếu không có kết quả
        allTranslatedTexts.push(translatedText);
        translatedCount++;
      }
      
      // Delay giữa các chunk
      const delay = selectedModel === 'gemma-3-4b-it' ? 2000 : 2500;
      if (chunkIndex < chunks.length - 1) {
        console.log(`⏳ Chờ ${delay}ms trước chunk tiếp theo...`);
        Utilities.sleep(delay);
      }
      
    } catch (error) {
      console.error(`❌ Lỗi chunk ${chunkIndex + 1}:`, error);
      errorCount += chunk.length;
      
      // Thêm text gốc vào kết quả nếu lỗi
      allTranslatedTexts.push(...chunk);
    }
  }
  
  // Cập nhật vùng với kết quả dịch
  console.log(`\n📝 Đang cập nhật vùng...`);
  
  const newValues = [];
  let textIndex = 0;
  
  for (let row = 0; row < numRows; row++) {
    const newRow = [];
    for (let col = 0; col < numCols; col++) {
      const cellValue = values[row][col];
      if (cellValue && typeof cellValue === 'string' && cellValue.trim() !== '') {
        if (textIndex < allTranslatedTexts.length) {
          const translatedText = allTranslatedTexts[textIndex];
          newRow.push(translatedText);
          console.log(`📝 Cập nhật ô ${row + 1},${col + 1}: "${cellValue}" -> "${translatedText}"`);
          textIndex++;
        } else {
          newRow.push(cellValue);
        }
      } else {
        newRow.push(cellValue);
      }
    }
    newValues.push(newRow);
  }
  
  // Cập nhật vùng
  range.setValues(newValues);
  
  const finalSuccessRate = textsToTranslate.length > 0 ? Math.round((translatedCount / textsToTranslate.length) * 100) : 0;
  
  console.log(`\n🎉 === KẾT QUẢ DỊCH VÙNG (TỐI ƯU) ===`);
  console.log(`📍 Vùng: ${range.getA1Notation()}`);
  console.log(`📊 Tổng số ô: ${totalCells}`);
  console.log(`📝 Số ô có nội dung: ${textsToTranslate.length}`);
  console.log(`✅ Số ô đã dịch: ${translatedCount}`);
  console.log(`❌ Số ô lỗi: ${errorCount}`);
  console.log(`⬜ Số ô rỗng: ${emptyCount}`);
  console.log(`📦 Số chunk đã xử lý: ${chunks.length}`);
  console.log(`📈 Tỷ lệ thành công: ${finalSuccessRate}%`);
  console.log(`🚀 Giảm API calls: ${Math.round(((textsToTranslate.length - chunks.length) / textsToTranslate.length) * 100)}%`);
  
  if (finalSuccessRate < 80) {
    console.log(`⚠️ Cảnh báo: Tỷ lệ thành công thấp (${finalSuccessRate}%). Có thể do:`);
    console.log(`   - Rate limit API`);
    console.log(`   - Quota exceeded`);
    console.log(`   - Network issues`);
    console.log(`   - API key không hợp lệ`);
  } else {
    console.log(`🎯 Tỷ lệ thành công tốt! (${finalSuccessRate}%)`);
  }
}