import i18n from 'i18next'
import { initReactI18next } from 'react-i18next'
import ru from './ru.json'
import en from './en.json'
import ky from './ky.json'
import vi from './vi.json'

const savedLang = localStorage.getItem('pebble_lang') || 'ru'

i18n
  .use(initReactI18next)
  .init({
    resources: {
      ru: { translation: ru },
      en: { translation: en },
      ky: { translation: ky },
      vi: { translation: vi },
    },
    lng: savedLang,
    fallbackLng: 'ru',
    interpolation: {
      escapeValue: false,
    },
  })

export default i18n

export const LANGUAGES = [
  { code: 'ru', label: 'RU', name: 'Русский' },
  { code: 'en', label: 'EN', name: 'English' },
  { code: 'ky', label: 'KY', name: 'Кыргызча' },
  { code: 'vi', label: 'VI', name: 'Tiếng Việt' },
] as const

export function changeLanguage(lang: string) {
  i18n.changeLanguage(lang)
  localStorage.setItem('pebble_lang', lang)
  // Dispatch event so components that fetch data can refetch with new lang
  window.dispatchEvent(new CustomEvent('pebble:langChange', { detail: { lang } }))
}

export function currentLang(): string {
  return i18n.language || 'ru'
}
