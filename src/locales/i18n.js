import { createI18n } from 'vue-i18n'
import zh from './zh.json'
import en from './en.json'
import jp from './jp.json'
import { bitable } from '@lark-base-open/js-sdk'

export const i18n = createI18n({
  locale: 'zh',
  fallbackLocale: 'en',
  allowComposition: true,
  messages: { zh, en, jp },
})

bitable.bridge.getLanguage().then((lang) => {
  const m = { zh: 'zh', 'zh-CN': 'zh', 'zh-TW': 'zh', en: 'en', 'en-US': 'en', jp: 'jp', ja: 'jp' }
  i18n.global.locale = m[lang] || 'en'
})