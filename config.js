/**
 * Configuration file for Google Sheets Translation Tool
 * File cấu hình cho tool dịch Google Sheets
 */

// ===== CẤU HÌNH CƠ BẢN =====

// API Configuration
const CONFIG = {
  // Gemini API settings
  GEMINI_API_URL: 'https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent',
  
  // Request settings
  REQUEST_DELAY: 100, // Delay giữa các request (milliseconds)
  MAX_RETRIES: 3,     // Số lần thử lại khi gặp lỗi
  TIMEOUT: 30000,     // Timeout cho mỗi request (milliseconds)
  
  // Batch processing
  BATCH_SIZE: 50,     // Số ô xử lý mỗi batch
  BATCH_DELAY: 1000,  // Delay giữa các batch (milliseconds)
  
  // UI settings
  PROGRESS_UPDATE_INTERVAL: 10, // Cập nhật progress mỗi N ô
  SHOW_DETAILED_PROGRESS: true, // Hiển thị progress chi tiết
};

// ===== NGÔN NGỮ HỖ TRỢ =====

const SUPPORTED_LANGUAGES = {
  '1': { code: 'Vietnamese', name: 'Tiếng Việt', native: 'Tiếng Việt' },
  '2': { code: 'English', name: 'Tiếng Anh', native: 'English' },
  '3': { code: 'Chinese', name: 'Tiếng Trung', native: '中文' },
  '4': { code: 'Japanese', name: 'Tiếng Nhật', native: '日本語' },
  '5': { code: 'Korean', name: 'Tiếng Hàn', native: '한국어' },
  '6': { code: 'French', name: 'Tiếng Pháp', native: 'Français' },
  '7': { code: 'German', name: 'Tiếng Đức', native: 'Deutsch' },
  '8': { code: 'Spanish', name: 'Tiếng Tây Ban Nha', native: 'Español' },
  '9': { code: 'Italian', name: 'Tiếng Ý', native: 'Italiano' },
  '10': { code: 'Portuguese', name: 'Tiếng Bồ Đào Nha', native: 'Português' },
  '11': { code: 'Russian', name: 'Tiếng Nga', native: 'Русский' },
  '12': { code: 'Arabic', name: 'Tiếng Ả Rập', native: 'العربية' },
  '13': { code: 'Hindi', name: 'Tiếng Hindi', native: 'हिन्दी' },
  '14': { code: 'Thai', name: 'Tiếng Thái', native: 'ไทย' },
  '15': { code: 'Indonesian', name: 'Tiếng Indonesia', native: 'Bahasa Indonesia' }
};

// ===== PROMPT TEMPLATES =====

const PROMPT_TEMPLATES = {
  // Template cơ bản
  basic: (text, targetLanguage) => 
    `Dịch đoạn text sau sang tiếng ${targetLanguage}. Chỉ trả về bản dịch, không thêm giải thích:
    
"${text}"`,

  // Template với context
  contextual: (text, targetLanguage, context) => 
    `Dịch đoạn text sau sang tiếng ${targetLanguage}. 
Context: ${context}
Chỉ trả về bản dịch, không thêm giải thích:

"${text}"`,

  // Template cho technical terms
  technical: (text, targetLanguage) => 
    `Dịch đoạn text sau sang tiếng ${targetLanguage}. 
Đây là thuật ngữ kỹ thuật, hãy dịch chính xác và phù hợp với ngữ cảnh.
Chỉ trả về bản dịch, không thêm giải thích:

"${text}"`,

  // Template cho formal language
  formal: (text, targetLanguage) => 
    `Dịch đoạn text sau sang tiếng ${targetLanguage}. 
Sử dụng ngôn ngữ trang trọng, lịch sự.
Chỉ trả về bản dịch, không thêm giải thích:

"${text}"`
};

// ===== ERROR MESSAGES =====

const ERROR_MESSAGES = {
  API_KEY_MISSING: 'API Key chưa được cài đặt. Vui lòng sử dụng menu "Cài đặt API Key"',
  API_KEY_INVALID: 'API Key không hợp lệ. Vui lòng kiểm tra lại',
  RATE_LIMIT: 'Đã vượt quá giới hạn request. Vui lòng thử lại sau',
  QUOTA_EXCEEDED: 'Đã vượt quá quota API. Vui lòng kiểm tra tài khoản Gemini',
  NETWORK_ERROR: 'Lỗi kết nối mạng. Vui lòng kiểm tra internet',
  TIMEOUT_ERROR: 'Request timeout. Vui lòng thử lại',
  INVALID_RESPONSE: 'Phản hồi không hợp lệ từ Gemini API',
  SHEET_EMPTY: 'Sheet không có dữ liệu để dịch',
  PERMISSION_DENIED: 'Không có quyền chỉnh sửa sheet này'
};

// ===== VALIDATION RULES =====

const VALIDATION_RULES = {
  // Kiểm tra text có cần dịch không
  shouldTranslate: (text) => {
    if (!text || text.toString().trim() === '') return false;
    
    // Không dịch số
    if (!isNaN(text) && text.toString().trim() !== '') return false;
    
    // Không dịch email
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (emailRegex.test(text)) return false;
    
    // Không dịch URL
    const urlRegex = /^https?:\/\/.+/;
    if (urlRegex.test(text)) return false;
    
    // Không dịch ngày tháng
    const dateRegex = /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/;
    if (dateRegex.test(text)) return false;
    
    return true;
  },
  
  // Kiểm tra API Key format
  isValidApiKey: (apiKey) => {
    return apiKey && apiKey.length > 20 && apiKey.startsWith('AIza');
  },
  
  // Kiểm tra ngôn ngữ hợp lệ
  isValidLanguage: (languageCode) => {
    return Object.values(SUPPORTED_LANGUAGES).some(lang => lang.code === languageCode);
  }
};

// ===== UTILITY FUNCTIONS =====

const UTILS = {
  // Format thời gian
  formatTime: (seconds) => {
    const mins = Math.floor(seconds / 60);
    const secs = seconds % 60;
    return `${mins}:${secs.toString().padStart(2, '0')}`;
  },
  
  // Format số với dấu phẩy
  formatNumber: (num) => {
    return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  },
  
  // Tạo ID duy nhất
  generateId: () => {
    return Math.random().toString(36).substr(2, 9);
  },
  
  // Delay function
  sleep: (ms) => {
    return new Promise(resolve => setTimeout(resolve, ms));
  },
  
  // Retry function
  retry: async (fn, maxRetries = CONFIG.MAX_RETRIES) => {
    for (let i = 0; i < maxRetries; i++) {
      try {
        return await fn();
      } catch (error) {
        if (i === maxRetries - 1) throw error;
        await UTILS.sleep(1000 * (i + 1)); // Exponential backoff
      }
    }
  }
};

// Export configuration
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    CONFIG,
    SUPPORTED_LANGUAGES,
    PROMPT_TEMPLATES,
    ERROR_MESSAGES,
    VALIDATION_RULES,
    UTILS
  };
} 