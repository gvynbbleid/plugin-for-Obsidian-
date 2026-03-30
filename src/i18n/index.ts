import deepmerge from 'deepmerge';
import en from './en';
import zh from './zh';

export default {
  i18n: {
    en,
    zh
  },
  get current() {
    const lang = window.localStorage.getItem('language') ?? 'en';
    return deepmerge(en, (this as any).i18n[lang as keyof typeof this.i18n] ?? {});
  }
};
